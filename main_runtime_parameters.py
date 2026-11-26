import argparse
import os
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from math import inf

import pandas as pd
from rapidfuzz import fuzz, process


# -------------------------------------------------------------------------
# Matching settings
# -------------------------------------------------------------------------

TOP_N = 3

# Less strict than the earlier word-pair version, while still filtering out
# weak description matches.
MIN_DESCRIPTION_SCORE = 50
MIN_PART_SCORE = 90

# The original fuzzy score remains the main description signal.
FUZZY_DESCRIPTION_WEIGHT = 0.95
WORD_PAIR_WEIGHT = 0.05

# Price is only allowed to reorder candidates whose description scores are
# close. It never increases the match score or rescues a poor description.
PRICE_TIE_WINDOW = 5.0

@dataclass(frozen=True)
class RecommendationSettings:
    """Runtime parameters for the recommendation engine.

    A future user interface can create this object from dropdowns, text fields,
    or branch-specific defaults and then pass it to run_recommendation().
    """

    # Normal user-facing options.
    preferred_suppliers: tuple[str, ...] = field(default_factory=tuple)
    excluded_suppliers: tuple[str, ...] = field(default_factory=tuple)
    in_stock_only: bool = False
    top_n: int = TOP_N

    # Advanced/admin options. Defaults preserve the current matching logic.
    min_description_score: float = MIN_DESCRIPTION_SCORE
    min_part_score: float = MIN_PART_SCORE
    fuzzy_description_weight: float = FUZZY_DESCRIPTION_WEIGHT
    word_pair_weight: float = WORD_PAIR_WEIGHT
    price_tie_window: float = PRICE_TIE_WINDOW
    description_candidate_multiplier: int = 10

    @property
    def description_candidate_limit(self) -> int:
        return self.top_n * self.description_candidate_multiplier

    def validate(self) -> None:
        if self.top_n < 1:
            raise ValueError("top_n must be at least 1.")

        for name, value in {
            "min_description_score": self.min_description_score,
            "min_part_score": self.min_part_score,
        }.items():
            if not 0 <= value <= 100:
                raise ValueError(f"{name} must be between 0 and 100.")

        for name, value in {
            "fuzzy_description_weight": self.fuzzy_description_weight,
            "word_pair_weight": self.word_pair_weight,
        }.items():
            if not 0 <= value <= 1:
                raise ValueError(f"{name} must be between 0 and 1.")

        weight_total = (
            self.fuzzy_description_weight
            + self.word_pair_weight
        )
        if abs(weight_total - 1.0) > 0.000001:
            raise ValueError(
                "fuzzy_description_weight and word_pair_weight "
                "must add up to 1.0."
            )

        if self.price_tie_window < 0:
            raise ValueError("price_tie_window cannot be negative.")

        if self.description_candidate_multiplier < 1:
            raise ValueError(
                "description_candidate_multiplier must be at least 1."
            )

OUTPUT_INVENTORY_COLUMNS = [
    "Product",
    "Group",
    "Description",
    "Supplier",
    "SupPart",
    "Unit",
    "CostEx",
    "ListPriceEx",
    "ListPriceInc",
    "OnHandQty",
]

# Description is the only mandatory customer field.
CUSTOMER_REQUIRED_COLUMNS = {"Description"}
CUSTOMER_OPTIONAL_COLUMNS = {
    "PartNumber": "",
    "CustomerPrice": pd.NA,
    "PriceBasis": "Unknown",
}

INVENTORY_REQUIRED_COLUMNS = {"SupPart", "Description"}

PRICE_BASIS_EX = "Ex GST"
PRICE_BASIS_INC = "Inc GST"
PRICE_BASIS_UNKNOWN = "Unknown"


def normalize_description(value: object) -> str:
    """Return a conservative normalized description for fuzzy matching."""
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKD", str(value))

    # Keep "non-permanent" distinct from "permanent". Without this, unrelated
    # non-permanent pencils can look similar to permanent markers.
    text = re.sub(
        r"\bnon[\s-]+permanent\b",
        "nonpermanent",
        text,
        flags=re.IGNORECASE,
    )

    text = text.replace("×", " x ")
    text = text.replace("&", " and ")
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.casefold()

    # Keep thousands together: 1,000 -> 1000.
    text = re.sub(r"(?<=\d),(?=\d{3}\b)", "", text)

    # Normalize dimensions: 100x200, 100 x 200 and 100X200.
    text = re.sub(r"(?<=\d)\s*x\s*(?=\d)", " x ", text)

    # Normalize joined letters and numbers: A4 -> A 4, 80GSM -> 80 GSM.
    text = re.sub(
        r"(?<=[a-z])(?=\d)|(?<=\d)(?=[a-z])",
        " ",
        text,
    )

    # Preserve decimal points only when they occur between digits.
    text = re.sub(r"(?<!\d)\.|\.(?!\d)", " ", text)
    text = re.sub(r"[^a-z0-9.]+", " ", text)

    return re.sub(r"\s+", " ", text).strip()


def normalize_part_number(value: object) -> str:
    """Normalize part numbers for exact and fuzzy matching."""
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value)).upper().strip()
    return re.sub(r"[^A-Z0-9]+", "", text)


def normalize_price(value: object) -> float | None:
    """Convert a numeric or currency-formatted value to a positive float."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        price = float(value)
        return price if price > 0 else None

    text = str(value).strip().replace(",", "")
    if not text:
        return None

    # Accept values such as "$45.00" or "NZD 45.00".
    cleaned = re.sub(r"[^0-9.\-]", "", text)

    try:
        price = float(cleaned)
    except (TypeError, ValueError):
        return None

    return price if price > 0 else None


def normalize_price_basis(value: object) -> str:
    """Normalize common ways of writing the customer price basis."""
    if pd.isna(value):
        return PRICE_BASIS_UNKNOWN

    text = re.sub(r"[^A-Z0-9]+", " ", str(value).upper()).strip()

    if text in {
        "EX",
        "EX GST",
        "EXCLUDING GST",
        "EXCLUDE GST",
        "GST EXCLUSIVE",
    }:
        return PRICE_BASIS_EX

    if text in {
        "INC",
        "INC GST",
        "INCLUDING GST",
        "INCLUDE GST",
        "GST INCLUSIVE",
    }:
        return PRICE_BASIS_INC

    return PRICE_BASIS_UNKNOWN


def build_word_pairs(description: str) -> list[tuple[str, str]]:
    """Return adjacent two-word pairs from a normalized description."""
    words = description.split()
    return list(zip(words, words[1:]))


def calculate_word_pair_score(
    customer_description: str,
    inventory_description: str,
) -> float:
    """Measure how many customer word pairs are represented by a candidate.

    The two words may appear in a different order or position in the
    inventory description. This keeps the validation layer light.
    """
    customer_pairs = build_word_pairs(customer_description)

    # A one-word description has no pair, so do not penalize it.
    if not customer_pairs:
        return 100.0

    inventory_words = set(inventory_description.split())

    matched_pairs = sum(
        1
        for first_word, second_word in customer_pairs
        if first_word in inventory_words and second_word in inventory_words
    )

    return (matched_pairs / len(customer_pairs)) * 100


def calculate_description_score(
    fuzzy_score: float,
    customer_description: str,
    inventory_description: str,
    settings: RecommendationSettings,
) -> float:
    """Combine the main RapidFuzz score with light word-pair validation."""
    word_pair_score = calculate_word_pair_score(
        customer_description,
        inventory_description,
    )

    return (
        fuzzy_score * settings.fuzzy_description_weight
        + word_pair_score * settings.word_pair_weight
    )


def normalize_supplier_name(value: object) -> str:
    """Normalize supplier names for case-insensitive exact filtering."""
    if pd.isna(value):
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).strip().casefold(),
    )


def get_available_suppliers(
    inventory_source: str | os.PathLike[str] | pd.DataFrame,
) -> list[str]:
    """Return cleaned supplier names for dropdowns or parameter selection."""
    if isinstance(inventory_source, pd.DataFrame):
        inventory_df = inventory_source
    else:
        inventory_df = pd.read_excel(inventory_source)

    if "Supplier" not in inventory_df.columns:
        return []

    suppliers = {
        str(value).strip()
        for value in inventory_df["Supplier"].dropna()
        if str(value).strip()
    }

    return sorted(suppliers, key=str.casefold)


def filter_inventory_by_suppliers(
    inventory_df: pd.DataFrame,
    preferred_suppliers: list[str],
    excluded_suppliers: list[str],
) -> pd.DataFrame:
    """Apply optional supplier inclusion and exclusion filters.

    An empty preferred-supplier list means all suppliers are initially allowed.
    Excluded suppliers are then removed. Exclusion takes priority when the same
    supplier appears in both lists.
    """
    normalized_preferred = {
        normalize_supplier_name(supplier)
        for supplier in preferred_suppliers
        if normalize_supplier_name(supplier)
    }
    normalized_excluded = {
        normalize_supplier_name(supplier)
        for supplier in excluded_suppliers
        if normalize_supplier_name(supplier)
    }

    if not normalized_preferred and not normalized_excluded:
        print(
            f"Supplier filter: not applied. "
            f"Searching all {len(inventory_df):,} inventory rows."
        )
        return inventory_df.copy()

    if "Supplier" not in inventory_df.columns:
        raise ValueError(
            "Supplier filters were provided, but the inventory file "
            "does not contain a 'Supplier' column."
        )

    supplier_normalized = inventory_df["Supplier"].map(
        normalize_supplier_name
    )

    supplier_mask = pd.Series(True, index=inventory_df.index, dtype=bool)

    if normalized_preferred:
        supplier_mask &= supplier_normalized.isin(normalized_preferred)

    if normalized_excluded:
        supplier_mask &= ~supplier_normalized.isin(normalized_excluded)

    filtered_df = inventory_df.loc[supplier_mask].copy()

    if filtered_df.empty:
        available_suppliers = sorted(
            {
                str(value).strip()
                for value in inventory_df["Supplier"].dropna()
                if str(value).strip()
            },
            key=str.casefold,
        )
        available_preview = ", ".join(available_suppliers[:20])

        raise ValueError(
            "The supplier filters returned no inventory rows. "
            f"Included: {', '.join(preferred_suppliers) or 'All suppliers'}. "
            f"Excluded: {', '.join(excluded_suppliers) or 'None'}. "
            f"Available supplier examples: {available_preview}"
        )

    include_summary = (
        ", ".join(preferred_suppliers)
        if preferred_suppliers
        else "All suppliers"
    )
    exclude_summary = (
        ", ".join(excluded_suppliers)
        if excluded_suppliers
        else "None"
    )

    print(
        f"Supplier filters applied: included={include_summary}; "
        f"excluded={exclude_summary}. "
        f"{len(filtered_df):,} of {len(inventory_df):,} inventory rows retained."
    )

    return filtered_df


def filter_inventory_by_stock(
    inventory_df: pd.DataFrame,
    in_stock_only: bool,
) -> pd.DataFrame:
    """Keep only inventory rows with a positive OnHandQty when requested."""
    if not in_stock_only:
        print(
            f"Stock filter: not applied. "
            f"Searching all {len(inventory_df):,} remaining inventory rows."
        )
        return inventory_df.copy()

    if "OnHandQty" not in inventory_df.columns:
        raise ValueError(
            "The in-stock-only filter was selected, but the inventory file "
            "does not contain an 'OnHandQty' column."
        )

    on_hand_quantity = pd.to_numeric(
        inventory_df["OnHandQty"],
        errors="coerce",
    ).fillna(0)

    filtered_df = inventory_df.loc[on_hand_quantity.gt(0)].copy()

    if filtered_df.empty:
        raise ValueError(
            "The in-stock-only filter returned no inventory rows with "
            "OnHandQty greater than zero."
        )

    print(
        f"Stock filter applied: OnHandQty > 0. "
        f"{len(filtered_df):,} of {len(inventory_df):,} inventory rows retained."
    )

    return filtered_df


def validate_columns(
    dataframe: pd.DataFrame,
    required: set[str],
    dataframe_name: str,
) -> None:
    missing = sorted(required.difference(dataframe.columns))
    if missing:
        raise ValueError(
            f"{dataframe_name} file is missing required column(s): "
            + ", ".join(missing)
        )


def add_optional_customer_columns(customer_df: pd.DataFrame) -> None:
    """Add optional template columns when an older customer file lacks them."""
    for column, default_value in CUSTOMER_OPTIONAL_COLUMNS.items():
        if column not in customer_df.columns:
            customer_df[column] = default_value


def add_optional_inventory_price_columns(inventory_df: pd.DataFrame) -> None:
    """Allow the script to run even if an inventory price column is absent."""
    for column in ("ListPriceEx", "ListPriceInc"):
        if column not in inventory_df.columns:
            inventory_df[column] = pd.NA
            print(
                f"Warning: inventory column '{column}' was not found. "
                "That price basis will not be available for matching."
            )


def build_exact_lookup(series: pd.Series) -> dict[str, list[int]]:
    """Map every non-empty normalized value to matching dataframe indices."""
    lookup: dict[str, list[int]] = defaultdict(list)

    for index, value in series.items():
        if value:
            lookup[value].append(index)

    return dict(lookup)


def extract_part_numbers_from_description(description: object) -> set[str]:
    """Extract conservative part-number candidates from customer text.

    This only supports exact embedded part-number matching. It deliberately
    does not fuzzy-match every number found inside a description.
    """
    if pd.isna(description):
        return set()

    raw_tokens = re.findall(
        r"[A-Za-z0-9]+(?:[-_./][A-Za-z0-9]+)*",
        str(description),
    )

    candidates: set[str] = set()

    for raw_token in raw_tokens:
        token = normalize_part_number(raw_token)

        if len(token) < 4 or not any(character.isdigit() for character in token):
            continue

        has_letter = any(character.isalpha() for character in token)
        is_long_numeric_code = token.isdigit() and len(token) >= 6

        if not has_letter and not is_long_numeric_code:
            continue

        # Ignore obvious measurements and pack quantities.
        if re.fullmatch(
            r"\d+(?:ML|L|G|GM|KG|MM|CM|M|PK|PACK|EA|EACH|PC|PCS)",
            token,
        ):
            continue

        candidates.add(token)

    return candidates


def get_price_comparison(
    customer_price: float | None,
    price_basis: str,
    inventory_row: pd.Series,
) -> dict[str, object]:
    """Compare customer price with the appropriate inventory list price."""
    empty_result = {
        "PriceBasisApplied": "",
        "PriceComparedAgainst": "",
        "InventoryComparisonPrice": None,
        "PriceDifference": None,
        "PriceDifferencePercent": None,
    }

    if customer_price is None:
        return empty_result

    list_price_ex = normalize_price(inventory_row.get("ListPriceEx_num"))
    list_price_inc = normalize_price(inventory_row.get("ListPriceInc_num"))

    possible_prices: list[tuple[str, float]] = []

    if price_basis == PRICE_BASIS_EX:
        if list_price_ex is not None:
            possible_prices.append(("ListPriceEx", list_price_ex))
        applied_basis = PRICE_BASIS_EX

    elif price_basis == PRICE_BASIS_INC:
        if list_price_inc is not None:
            possible_prices.append(("ListPriceInc", list_price_inc))
        applied_basis = PRICE_BASIS_INC

    else:
        if list_price_ex is not None:
            possible_prices.append(("ListPriceEx", list_price_ex))
        if list_price_inc is not None:
            possible_prices.append(("ListPriceInc", list_price_inc))
        applied_basis = "Unknown - closest Ex/Inc price used"

    if not possible_prices:
        return {
            **empty_result,
            "PriceBasisApplied": applied_basis,
        }

    compared_column, inventory_price = min(
        possible_prices,
        key=lambda item: abs(item[1] - customer_price),
    )

    price_difference = abs(inventory_price - customer_price)
    price_difference_percent = (
        price_difference / customer_price
    ) * 100

    return {
        "PriceBasisApplied": applied_basis,
        "PriceComparedAgainst": compared_column,
        "InventoryComparisonPrice": inventory_price,
        "PriceDifference": price_difference,
        "PriceDifferencePercent": price_difference_percent,
    }


def sort_exact_matches_by_price(
    inventory_indices: list[int],
    inventory_df: pd.DataFrame,
    customer_price: float | None,
    price_basis: str,
) -> list[int]:
    """Use closest price only to order duplicate exact part-number rows."""
    if customer_price is None or len(inventory_indices) < 2:
        return inventory_indices

    def price_sort_key(inventory_index: int) -> tuple[bool, float]:
        comparison = get_price_comparison(
            customer_price,
            price_basis,
            inventory_df.loc[inventory_index],
        )
        difference_percent = comparison["PriceDifferencePercent"]

        return (
            difference_percent is None,
            difference_percent if difference_percent is not None else inf,
        )

    return sorted(inventory_indices, key=price_sort_key)


def sort_description_candidates(
    candidates: list[dict[str, object]],
    customer_price: float | None,
    settings: RecommendationSettings,
) -> list[dict[str, object]]:
    """Rank description candidates by score from highest to lowest.

    Match score is always the primary ranking rule. Customer price is used
    only as a tie-breaker when two candidates have exactly the same underlying
    score. It can never move a lower-scoring product above a higher-scoring
    product.

    The settings argument is retained for compatibility with the existing
    function call.
    """
    del settings

    def candidate_sort_key(
        candidate: dict[str, object],
    ) -> tuple[float, bool, float]:
        score = float(candidate["score"])

        if customer_price is None:
            return (-score, False, 0.0)

        difference_percent = candidate["price"][
            "PriceDifferencePercent"
        ]

        return (
            -score,
            difference_percent is None,
            (
                float(difference_percent)
                if difference_percent is not None
                else inf
            ),
        )

    return sorted(
        candidates,
        key=candidate_sort_key,
    )


def add_output(
    rows: list[dict],
    customer_row: pd.Series,
    inventory_row: pd.Series,
    rank: int,
    score: float,
    match_type: str,
    price_comparison: dict[str, object],
) -> None:
    """Add one recommendation while keeping helper columns internal."""
    output = {
        f"Customer_{column}": customer_row[column]
        for column in customer_row.index
        if not column.endswith("_norm")
        and not column.endswith("_parsed")
    }

    output["Rank"] = rank
    output["Score"] = round(score)
    output["MatchType"] = match_type

    for column in OUTPUT_INVENTORY_COLUMNS:
        if column in inventory_row.index:
            output[f"Inventory_{column}"] = inventory_row[column]

    output["PriceBasisApplied"] = price_comparison["PriceBasisApplied"]
    output["PriceComparedAgainst"] = price_comparison[
        "PriceComparedAgainst"
    ]
    output["InventoryComparisonPrice"] = price_comparison[
        "InventoryComparisonPrice"
    ]
    output["PriceDifference"] = price_comparison["PriceDifference"]
    output["PriceDifferencePercent"] = price_comparison[
        "PriceDifferencePercent"
    ]

    rows.append(output)


def add_no_match_output(
    rows: list[dict],
    customer_row: pd.Series,
) -> None:
    """Preserve an unmatched enquiry with blank recommendation fields."""
    output = {
        f"Customer_{column}": customer_row[column]
        for column in customer_row.index
        if not column.endswith("_norm")
        and not column.endswith("_parsed")
    }

    output["Rank"] = pd.NA
    output["Score"] = pd.NA
    output["MatchType"] = ""

    for column in OUTPUT_INVENTORY_COLUMNS:
        output[f"Inventory_{column}"] = pd.NA

    output["PriceBasisApplied"] = ""
    output["PriceComparedAgainst"] = ""
    output["InventoryComparisonPrice"] = pd.NA
    output["PriceDifference"] = pd.NA
    output["PriceDifferencePercent"] = pd.NA

    rows.append(output)


def run_recommendation(
    customer_file: str | os.PathLike[str],
    inventory_file: str | os.PathLike[str],
    output_file: str | os.PathLike[str] = os.path.join(
        "output",
        "recommendations.xlsx",
    ),
    settings: RecommendationSettings | None = None,
) -> pd.DataFrame:
    """Run the recommendation process and return the output dataframe.

    This function is the integration point for a future Power Apps, web,
    desktop, or other user interface. The interface only needs to collect the
    file locations and settings, then call this function.
    """
    settings = settings or RecommendationSettings()
    settings.validate()

    customer_df = pd.read_excel(customer_file)
    inventory_df = pd.read_excel(inventory_file)

    validate_columns(
        customer_df,
        CUSTOMER_REQUIRED_COLUMNS,
        "Customer",
    )
    validate_columns(
        inventory_df,
        INVENTORY_REQUIRED_COLUMNS,
        "Inventory",
    )

    add_optional_customer_columns(customer_df)
    add_optional_inventory_price_columns(inventory_df)

    if "ObsoleteFlag" in inventory_df.columns:
        obsolete_flag = (
            inventory_df["ObsoleteFlag"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        inventory_df = inventory_df.loc[obsolete_flag.ne("Y")].copy()

    # Runtime inventory filters are applied before matching lookups are built.
    inventory_df = filter_inventory_by_suppliers(
        inventory_df,
        list(settings.preferred_suppliers),
        list(settings.excluded_suppliers),
    )
    inventory_df = filter_inventory_by_stock(
        inventory_df,
        settings.in_stock_only,
    )

    inventory_df.reset_index(drop=True, inplace=True)

    # Temporary normalized/helper columns used only during matching.
    customer_df["PartNumber_norm"] = customer_df["PartNumber"].map(
        normalize_part_number
    )
    customer_df["Description_norm"] = customer_df["Description"].map(
        normalize_description
    )
    customer_df["CustomerPrice_parsed"] = customer_df[
        "CustomerPrice"
    ].map(normalize_price)
    customer_df["PriceBasis_norm"] = customer_df["PriceBasis"].map(
        normalize_price_basis
    )

    inventory_df["SupPart_norm"] = inventory_df["SupPart"].map(
        normalize_part_number
    )
    inventory_df["Description_norm"] = inventory_df["Description"].map(
        normalize_description
    )
    inventory_df["ListPriceEx_num"] = inventory_df["ListPriceEx"].map(
        normalize_price
    )
    inventory_df["ListPriceInc_num"] = inventory_df["ListPriceInc"].map(
        normalize_price
    )

    exact_part_lookup = build_exact_lookup(
        inventory_df["SupPart_norm"]
    )

    description_choices = {
        index: value
        for index, value in inventory_df["Description_norm"].items()
        if value
    }
    part_choices = {
        index: value
        for index, value in inventory_df["SupPart_norm"].items()
        if value
    }

    rows: list[dict] = []

    for _, customer_row in customer_df.iterrows():
        description = customer_row["Description_norm"]

        # Skip completely blank enquiry rows from the template.
        if not description:
            continue

        enquiry_output_start = len(rows)
        used_inventory_indices: set[int] = set()
        rank = 1

        part_number = customer_row["PartNumber_norm"]
        customer_price = customer_row["CustomerPrice_parsed"]
        price_basis = customer_row["PriceBasis_norm"]

        def append_match(
            inventory_index: int,
            score: float,
            match_type: str,
        ) -> None:
            nonlocal rank

            price_comparison = get_price_comparison(
                customer_price,
                price_basis,
                inventory_df.loc[inventory_index],
            )

            add_output(
                rows,
                customer_row,
                inventory_df.loc[inventory_index],
                rank,
                score,
                match_type,
                price_comparison,
            )

            used_inventory_indices.add(inventory_index)
            rank += 1

        # 1. Exact part number entered in the PartNumber column.
        if part_number:
            exact_indices = exact_part_lookup.get(part_number, [])
            exact_indices = sort_exact_matches_by_price(
                exact_indices,
                inventory_df,
                customer_price,
                price_basis,
            )

            for inventory_index in exact_indices:
                if rank > settings.top_n:
                    break

                append_match(
                    inventory_index,
                    100,
                    "Exact Part Number",
                )

        # 2. Exact inventory part number found inside customer Description.
        if rank <= settings.top_n:
            embedded_part_numbers = extract_part_numbers_from_description(
                customer_row["Description"]
            )

            # Also treat the complete Description value as a possible part
            # number. This supports direct searches for short, letters-only,
            # numeric, dashed, or slashed part numbers that the conservative
            # embedded-token extractor may intentionally ignore.
            direct_description_part_number = normalize_part_number(
                customer_row["Description"]
            )
            if direct_description_part_number in exact_part_lookup:
                embedded_part_numbers.add(
                    direct_description_part_number
                )

            embedded_indices: list[int] = []

            for embedded_part_number in embedded_part_numbers:
                for inventory_index in exact_part_lookup.get(
                    embedded_part_number,
                    [],
                ):
                    if inventory_index not in used_inventory_indices:
                        embedded_indices.append(inventory_index)

            embedded_indices = list(dict.fromkeys(embedded_indices))
            embedded_indices = sort_exact_matches_by_price(
                embedded_indices,
                inventory_df,
                customer_price,
                price_basis,
            )

            for inventory_index in embedded_indices:
                if rank > settings.top_n:
                    break

                append_match(
                    inventory_index,
                    100,
                    "Exact Part Number in Description",
                )

        # 3. Fuzzy part-number matches from the explicit PartNumber column.
        if part_number and rank <= settings.top_n:
            part_matches = process.extract(
                part_number,
                part_choices,
                scorer=fuzz.ratio,
                score_cutoff=settings.min_part_score,
                limit=settings.top_n * 3,
            )

            for _, score, inventory_index in part_matches:
                if rank > settings.top_n:
                    break
                if inventory_index in used_inventory_indices:
                    continue

                append_match(
                    inventory_index,
                    score,
                    "Fuzzy Part Number",
                )

        # 4. Description matching fills any ranks still available.
        if description and rank <= settings.top_n:
            # token_sort_ratio is more literal than WRatio. It handles word
            # reordering without inflating partial or contained descriptions.
            #
            # The cutoff below is only a retrieval optimisation. The sole
            # acceptance threshold remains settings.min_description_score.
            if settings.fuzzy_description_weight > 0:
                maximum_word_pair_contribution = (
                    settings.word_pair_weight * 100
                )
                retrieval_cutoff = max(
                    0.0,
                    (
                        settings.min_description_score
                        - maximum_word_pair_contribution
                    )
                    / settings.fuzzy_description_weight,
                )
            else:
                retrieval_cutoff = 0.0

            description_matches = process.extract(
                description,
                description_choices,
                scorer=fuzz.token_sort_ratio,
                score_cutoff=retrieval_cutoff,
                limit=settings.description_candidate_limit,
            )

            description_candidates: list[dict[str, object]] = []

            for _, fuzzy_score, inventory_index in description_matches:
                if inventory_index in used_inventory_indices:
                    continue

                inventory_description = inventory_df.at[
                    inventory_index,
                    "Description_norm",
                ]

                final_score = calculate_description_score(
                    fuzzy_score,
                    description,
                    inventory_description,
                    settings,
                )

                # This is the only description-match acceptance threshold.
                if final_score < settings.min_description_score:
                    continue

                price_comparison = get_price_comparison(
                    customer_price,
                    price_basis,
                    inventory_df.loc[inventory_index],
                )

                description_candidates.append(
                    {
                        "inventory_index": inventory_index,
                        "score": final_score,
                        "price": price_comparison,
                    }
                )

            description_candidates = sort_description_candidates(
                description_candidates,
                customer_price,
                settings,
            )

            for candidate in description_candidates:
                if rank > settings.top_n:
                    break

                inventory_index = int(candidate["inventory_index"])

                if inventory_index in used_inventory_indices:
                    continue

                add_output(
                    rows,
                    customer_row,
                    inventory_df.loc[inventory_index],
                    rank,
                    float(candidate["score"]),
                    "Description + Word Pairs",
                    candidate["price"],
                )

                used_inventory_indices.add(inventory_index)
                rank += 1

        if len(rows) == enquiry_output_start:
            add_no_match_output(
                rows,
                customer_row,
            )

    output_path = os.fspath(output_file)
    output_directory = os.path.dirname(output_path) or "."
    os.makedirs(output_directory, exist_ok=True)

    output_df = pd.DataFrame(rows)
    output_df.to_excel(
        output_path,
        index=False,
        sheet_name="Recommendations",
    )

    supplier_summary = (
        ", ".join(settings.preferred_suppliers)
        if settings.preferred_suppliers
        else "All suppliers"
    )
    excluded_supplier_summary = (
        ", ".join(settings.excluded_suppliers)
        if settings.excluded_suppliers
        else "None"
    )

    matched_recommendation_count = (
        output_df["MatchType"]
        .fillna("")
        .astype(str)
        .str.strip()
        .ne("")
        .sum()
        if "MatchType" in output_df.columns
        else 0
    )
    unmatched_enquiry_count = (
        output_df["MatchType"]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("")
        .sum()
        if "MatchType" in output_df.columns
        else 0
    )

    print(
        f"Done. {matched_recommendation_count:,} matched recommendations "
        f"and {unmatched_enquiry_count:,} unmatched enquiry rows saved "
        f"to {output_path}"
    )
    print(
        f"Parameters used: suppliers={supplier_summary}; "
        f"excluded_suppliers={excluded_supplier_summary}; "
        f"in_stock_only={settings.in_stock_only}; "
        f"top_n={settings.top_n}; "
        f"min_description_score={settings.min_description_score}; "
        f"description_scorer=token_sort_ratio; "
        f"min_part_score={settings.min_part_score}"
    )

    return output_df


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Recommend inventory products using part-number, description, "
            "optional customer-price matching, and runtime supplier filters."
        )
    )

    parser.add_argument(
        "--customer",
        help="Customer enquiry Excel input file",
    )
    parser.add_argument(
        "--inventory",
        required=True,
        help="Inventory Excel input file",
    )
    parser.add_argument(
        "--output",
        default=os.path.join("output", "recommendations.xlsx"),
        help="Output Excel file path",
    )

    # Main user-facing controls.
    parser.add_argument(
        "--preferred-supplier",
        action="append",
        default=[],
        help=(
            "Restrict inventory to a supplier. Repeat this option for more "
            "than one supplier. Omit it to search all suppliers."
        ),
    )
    parser.add_argument(
        "--exclude-supplier",
        action="append",
        default=[],
        help=(
            "Exclude a supplier from inventory matching. Repeat this option "
            "for more than one supplier. Exclusion takes priority over inclusion."
        ),
    )
    parser.add_argument(
        "--in-stock-only",
        action="store_true",
        help=(
            "Search only inventory rows where OnHandQty is greater than zero. "
            "The default is to include all stock quantities."
        ),
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=TOP_N,
        help=f"Maximum results per enquiry. Default: {TOP_N}",
    )
    parser.add_argument(
        "--list-suppliers",
        action="store_true",
        help=(
            "Print supplier names available in the inventory file and exit. "
            "Useful for populating a future dropdown."
        ),
    )

    # Advanced/admin controls. Normal sales users should usually keep defaults.
    advanced = parser.add_argument_group("advanced matching parameters")
    advanced.add_argument(
        "--min-description-score",
        type=float,
        default=MIN_DESCRIPTION_SCORE,
    )
    advanced.add_argument(
        "--min-part-score",
        type=float,
        default=MIN_PART_SCORE,
    )
    advanced.add_argument(
        "--fuzzy-description-weight",
        type=float,
        default=FUZZY_DESCRIPTION_WEIGHT,
    )
    advanced.add_argument(
        "--word-pair-weight",
        type=float,
        default=WORD_PAIR_WEIGHT,
    )
    advanced.add_argument(
        "--price-tie-window",
        type=float,
        default=PRICE_TIE_WINDOW,
    )

    return parser


def settings_from_arguments(
    args: argparse.Namespace,
) -> RecommendationSettings:
    return RecommendationSettings(
        preferred_suppliers=tuple(args.preferred_supplier),
        excluded_suppliers=tuple(args.exclude_supplier),
        in_stock_only=args.in_stock_only,
        top_n=args.top_n,
        min_description_score=args.min_description_score,
        min_part_score=args.min_part_score,
        fuzzy_description_weight=args.fuzzy_description_weight,
        word_pair_weight=args.word_pair_weight,
        price_tie_window=args.price_tie_window,
    )


def main() -> None:
    parser = build_argument_parser()
    args = parser.parse_args()

    if args.list_suppliers:
        suppliers = get_available_suppliers(args.inventory)

        if not suppliers:
            print("No Supplier column or supplier values were found.")
            return

        print("Available suppliers:")
        for supplier in suppliers:
            print(f"  - {supplier}")
        return

    if not args.customer:
        parser.error(
            "--customer is required unless --list-suppliers is used."
        )

    settings = settings_from_arguments(args)

    run_recommendation(
        customer_file=args.customer,
        inventory_file=args.inventory,
        output_file=args.output,
        settings=settings,
    )


if __name__ == "__main__":
    main()