from __future__ import annotations

from datetime import datetime
from io import BytesIO
import re
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from main_runtime_parameters import (
    MIN_DESCRIPTION_SCORE,
    RecommendationSettings,
    get_available_suppliers,
    normalize_part_number,
    run_recommendation,
)


APP_TITLE = "OPD Product Recommendation Tool"
INVENTORY_FILE = Path("data") / "inventory" / "inventory.xlsx"


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🔎",
    layout="centered",
)



# -------------------------------------------------------------------------
# Application styling
# -------------------------------------------------------------------------

COMPANY_ORANGE = "#F47927"
DARK_GRAY = "#333333"
MEDIUM_GRAY = "#666666"
LIGHT_GRAY = "#E6E6E6"

st.markdown(
    f"""
    <style>
        :root {{
            --company-orange: {COMPANY_ORANGE};
            --company-dark-gray: {DARK_GRAY};
            --company-medium-gray: {MEDIUM_GRAY};
            --company-light-gray: {LIGHT_GRAY};
        }}

        .stApp {{
            background-color: #FFFFFF;
            color: var(--company-dark-gray);
        }}

        .block-container {{
            max-width: 1100px;
            padding-top: 2rem;
            padding-bottom: 3rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }}

        
        h1 {{
            color: var(--company-orange);
            border-bottom: 4px solid var(--company-orange);
            padding-bottom: 0.55rem;
            margin-bottom: 0.4rem;
            font-size: 1.5rem;
        }}

        h2, h3 {{
            color: var(--company-orange);
        }}

        /* Override only the main Streamlit title */
        [data-testid="stHeadingWithActionElements"] h1,
        [data-testid="stHeadingWithActionElements"] h1 span {{
            color: var(--company-orange) !important;
            font-size: 2rem !important;
            border-bottom: none !important;
            padding-top: 12px !important;            
        }}

        p, label, .stCaption {{
            color: var(--company-medium-gray);
        }}

        a {{
            color: var(--company-orange) !important;
        }}

        .stButton > button,
        .stDownloadButton > button {{
            border-radius: 8px;
            border: 1px solid var(--company-orange);
            font-weight: 600;
            min-height: 2.9rem;
        }}

        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {{
            background-color: var(--company-orange) !important;
            border-color: var(--company-orange) !important;
            color: #FFFFFF !important;
        }}

        .stButton > button[kind="primary"] p,
        .stButton > button[kind="primary"] span,
        .stDownloadButton > button[kind="primary"] p,
        .stDownloadButton > button[kind="primary"] span {{
            color: #FFFFFF !important;
        }}

        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover {{
            background-color: #D96518;
            border-color: #D96518;
            color: #FFFFFF;
        }}

        .stButton > button:not([kind="primary"]),
        .stDownloadButton > button:not([kind="primary"]) {{
            background-color: #FFFFFF;
            color: var(--company-orange);
        }}

        .stButton > button:not([kind="primary"]):hover,
        .stDownloadButton > button:not([kind="primary"]):hover {{
            background-color: #FFF4EC;
            color: #D96518;
            border-color: #D96518;
        }}

        input:focus,
        textarea:focus,
        select:focus {{
            border-color: var(--company-orange) !important;
            box-shadow: 0 0 0 1px var(--company-orange) !important;
        }}

        [data-testid="stMetric"] {{
            background-color: #FFFFFF;
            border: 1px solid var(--company-light-gray);
            border-top: 4px solid var(--company-orange);
            border-radius: 10px;
            padding: 1rem 1.1rem;
        }}

        [data-testid="stMetricValue"] {{
            color: var(--company-dark-gray);
            font-size: 1.5rem;
        }}

        [data-testid="stMetricLabel"] {{
            color: var(--company-medium-gray);
        }}

        [data-testid="stFileUploaderDropzone"] {{
            background-color: #FAFAFA;
            border: 1px dashed #BDBDBD;
            border-radius: 10px;
        }}

        [data-testid="stFileUploaderDropzone"]:hover {{
            border-color: var(--company-orange);
            background-color: #FFF8F3;
        }}

        [data-testid="stFileUploader"] button {{
            background-color: #F47927 !important;
            color: #FFFFFF !important;
            border: 1px solid #F47927 !important;
            border-radius: 8px !important;
        }}      

        hr {{
            border-color: var(--company-light-gray);
        }}

        [data-testid="stExpander"] {{
            border: 1px solid var(--company-light-gray);
            border-radius: 10px;
            background-color: #FAFAFA;
        }}     

        [data-testid="stDataFrame"] {{
            background-color: #FFFFFF !important;
            border: 1px solid var(--company-light-gray);
            border-radius: 10px;
            overflow: hidden;
        }}

        [data-testid="stDataFrame"] > div {{
            background-color: #FFFFFF !important;
        }}

        [data-testid="stDataFrame"] canvas {{
            background-color: #FFFFFF !important;
        }}


        /* Product search assistant popup */
        [role="dialog"] {{
            background-color: #FFFFFF !important;
            border-radius: 12px !important;
        }}

        [role="dialog"] > div {{
            background-color: #FFFFFF !important;
        }}

        [role="dialog"] h2 {{
            color: var(--company-orange) !important;
        }}

        [data-testid="stChatMessage"] {{
            background-color: #FFFFFF;
            border: 1px solid var(--company-light-gray);
            border-radius: 10px;
            padding: 0.65rem 0.9rem;
            margin-bottom: 0.75rem;
        }}

        [data-testid="stChatInput"] {{
            background-color: #FFFFFF;
            border-top: 1px solid var(--company-light-gray);
        }}

        [data-testid="stChatInput"] textarea {{
            background-color: #F7F7F7 !important;
            color: var(--company-dark-gray) !important;
        }}

        [data-testid="stChatInput"] button {{
            background-color: var(--company-orange) !important;
            color: #FFFFFF !important;
            border-radius: 8px !important;
        }}

        @media (max-width: 800px) {{
            .block-container {{
                padding-left: 1rem;
                padding-right: 1rem;
                padding-top: 1rem;
            }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data(show_spinner=False)
def generate_customer_enquiry_template() -> bytes:
    """Create the customer enquiry Excel template in memory."""
    workbook = Workbook()

    enquiry_sheet = workbook.active
    enquiry_sheet.title = "Customer Enquiries"

    headers = [
        "Description",
        "PartNumber",
        "CustomerPrice",
        "PriceBasis",
    ]
    enquiry_sheet.append(headers)

    company_orange = "F47927"
    white = "FFFFFF"
    light_gray = "F2F2F2"
    dark_gray = "333333"

    header_fill = PatternFill(
        fill_type="solid",
        fgColor=company_orange,
    )
    header_font = Font(
        color=white,
        bold=True,
    )

    for cell in enquiry_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    enquiry_sheet.freeze_panes = "A2"
    enquiry_sheet.auto_filter.ref = "A1:D101"

    enquiry_sheet.column_dimensions["A"].width = 52
    enquiry_sheet.column_dimensions["B"].width = 22
    enquiry_sheet.column_dimensions["C"].width = 18
    enquiry_sheet.column_dimensions["D"].width = 18
    enquiry_sheet.row_dimensions[1].height = 24

    # Add blank entry rows.
    for row_number in range(2, 102):
        enquiry_sheet.cell(row=row_number, column=1, value="")
        enquiry_sheet.cell(row=row_number, column=2, value="")
        enquiry_sheet.cell(row=row_number, column=3, value=None)
        enquiry_sheet.cell(row=row_number, column=4, value="")

        enquiry_sheet.cell(
            row=row_number,
            column=3,
        ).number_format = '$#,##0.00'

    # Price-basis dropdown.
    price_basis_validation = DataValidation(
        type="list",
        formula1='"Ex GST,Inc GST,Unknown"',
        allow_blank=True,
    )
    price_basis_validation.error = (
        "Select Ex GST, Inc GST, or Unknown."
    )
    price_basis_validation.errorTitle = "Invalid price basis"
    price_basis_validation.prompt = (
        "Choose whether the customer price includes GST."
    )
    price_basis_validation.promptTitle = "Price basis"
    enquiry_sheet.add_data_validation(price_basis_validation)
    price_basis_validation.add("D2:D101")

    # Instructions sheet.
    instructions_sheet = workbook.create_sheet("Instructions")
    instructions_sheet["A1"] = "Customer Enquiry Template"
    instructions_sheet["A1"].fill = header_fill
    instructions_sheet["A1"].font = Font(
        color=white,
        bold=True,
        size=14,
    )
    instructions_sheet["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )
    instructions_sheet.merge_cells("A1:B1")
    instructions_sheet.row_dimensions[1].height = 28

    instructions = [
        (
            "Description",
            "Required. Enter the customer's product wording, including "
            "brand, colour, size, unit, pack quantity, or any part number "
            "mentioned in the enquiry.",
        ),
        (
            "PartNumber",
            "Optional. Enter a known supplier or customer part number.",
        ),
        (
            "CustomerPrice",
            "Optional. Enter the price mentioned by the customer. Price is "
            "used only to help rank closely matched products or unit variants.",
        ),
        (
            "PriceBasis",
            "Optional. Select Ex GST, Inc GST, or Unknown.",
        ),
        (
            "Important",
            "Enter one customer-requested item per row. Unused blank rows "
            "are ignored by the recommendation tool.",
        ),
    ]

    instructions_sheet["A2"] = "Column"
    instructions_sheet["B2"] = "How to use it"

    for cell in instructions_sheet[2]:
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=light_gray,
        )
        cell.font = Font(
            color=dark_gray,
            bold=True,
        )

    for row_number, (column_name, guidance) in enumerate(
        instructions,
        start=3,
    ):
        instructions_sheet.cell(
            row=row_number,
            column=1,
            value=column_name,
        )
        instructions_sheet.cell(
            row=row_number,
            column=2,
            value=guidance,
        )

        instructions_sheet.cell(
            row=row_number,
            column=1,
        ).font = Font(
            color=dark_gray,
            bold=True,
        )
        instructions_sheet.cell(
            row=row_number,
            column=2,
        ).alignment = Alignment(
            wrap_text=True,
            vertical="top",
        )

    instructions_sheet.column_dimensions["A"].width = 22
    instructions_sheet.column_dimensions["B"].width = 85
    instructions_sheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()


@st.cache_data(show_spinner=False)
def load_inventory_summary(
    inventory_path: str,
    modified_time: float,
) -> tuple[list[str], int, set[str]]:
    """Load supplier choices, row count, and normalized part numbers."""
    del modified_time  # Used only to invalidate the Streamlit cache.
    inventory_df = pd.read_excel(inventory_path)
    suppliers = get_available_suppliers(inventory_df)

    inventory_part_numbers: set[str] = set()
    if "SupPart" in inventory_df.columns:
        inventory_part_numbers = {
            normalized_part
            for normalized_part in inventory_df["SupPart"].map(
                normalize_part_number
            )
            if normalized_part
        }

    return suppliers, len(inventory_df), inventory_part_numbers


def clear_previous_results() -> None:
    st.session_state.pop("recommendation_bytes", None)
    st.session_state.pop("recommendation_preview", None)
    st.session_state.pop("recommendation_filename", None)
    st.session_state.pop("recommendation_metrics", None)


def parse_chat_descriptions(prompt: str) -> list[str]:
    """Return non-empty product descriptions entered one per line."""
    descriptions = [
        line.strip()
        for line in prompt.splitlines()
        if line.strip()
    ]

    # Preserve order while removing duplicate lines.
    return list(dict.fromkeys(descriptions))


def looks_like_part_number_query(
    query: str,
    inventory_part_numbers: set[str],
) -> bool:
    """Return True when a chat line should also use PartNumber matching."""
    normalized_query = normalize_part_number(query)

    if not normalized_query:
        return False

    # Strongest signal: the complete query exactly matches an inventory part
    # number after normalizing case, spaces, dashes, and punctuation.
    if normalized_query in inventory_part_numbers:
        return True

    stripped_query = query.strip()
    query_words = stripped_query.split()

    # A compact single-token identifier containing digits or common part-code
    # separators should also use fuzzy PartNumber matching.
    return (
        len(query_words) == 1
        and len(normalized_query) >= 3
        and (
            any(character.isdigit() for character in normalized_query)
            or bool(re.search(r"[-_./]", stripped_query))
        )
    )


def build_chat_customer_workbook(
    descriptions: list[str],
    output_path: Path,
    inventory_part_numbers: set[str],
) -> None:
    """Create a temporary enquiry workbook for description or part searches."""
    detected_part_numbers = [
        description
        if looks_like_part_number_query(
            description,
            inventory_part_numbers,
        )
        else ""
        for description in descriptions
    ]

    customer_df = pd.DataFrame(
        {
            # Description remains populated because it is the required field
            # and provides a fallback when the query is not a part number.
            "Description": descriptions,
            "PartNumber": detected_part_numbers,
            "CustomerPrice": [pd.NA] * len(descriptions),
            "PriceBasis": ["Unknown"] * len(descriptions),
        }
    )
    customer_df.to_excel(
        output_path,
        index=False,
        sheet_name="Customer Enquiries",
    )


def get_chat_result_columns(
    recommendation_df: pd.DataFrame,
) -> list[str]:
    """Return the concise columns shown inside chatbot responses."""
    return [
        column
        for column in [
            "Customer_Description",
            "Rank",
            "Score",
            "Inventory_Product",
            "Inventory_Description",
            "Inventory_Supplier",
            "Inventory_SupPart",
            "Inventory_Unit",
            "Inventory_ListPriceEx",
            "Inventory_ListPriceInc",
            "Inventory_OnHandQty",
            "MatchType"
        ]
        if column in recommendation_df.columns
    ]


def run_chat_product_search(
    prompt: str,
    settings: RecommendationSettings,
    inventory_part_numbers: set[str],
) -> tuple[pd.DataFrame, bytes, list[str]]:
    """Run the existing recommendation engine for chat-entered descriptions."""
    descriptions = parse_chat_descriptions(prompt)

    if not descriptions:
        raise ValueError(
            "Enter at least one product description."
        )

    with TemporaryDirectory() as temp_directory:
        temp_path = Path(temp_directory)
        customer_path = temp_path / "chat_customer_enquiry.xlsx"
        output_path = temp_path / "chat_recommendations.xlsx"

        build_chat_customer_workbook(
            descriptions,
            customer_path,
            inventory_part_numbers,
        )

        recommendation_df = run_recommendation(
            customer_file=customer_path,
            inventory_file=INVENTORY_FILE,
            output_file=output_path,
            settings=settings,
        )

        output_bytes = output_path.read_bytes()

    return recommendation_df, output_bytes, descriptions


@st.dialog(
    "🤖 Product Search Assistant",
    width="large",
)
def show_product_search_assistant(
    initial_settings: RecommendationSettings,
    available_suppliers: list[str],
    inventory_part_numbers: set[str],
) -> None:
    """Display the local product-search assistant in a modal dialog."""
    if "product_chat_messages" not in st.session_state:
        st.session_state["product_chat_messages"] = []

    # Initialise chatbot parameters from the main page only when the popup
    # is first opened. After that, the chatbot controls are independent.
    chat_parameter_defaults = {
        "chat_search_all_suppliers": (
            not bool(initial_settings.preferred_suppliers)
        ),
        "chat_preferred_suppliers": list(
            initial_settings.preferred_suppliers
        ),
        "chat_excluded_suppliers": list(
            initial_settings.excluded_suppliers
        ),
        "chat_in_stock_only": initial_settings.in_stock_only,
        "chat_top_n": initial_settings.top_n,
        "chat_min_description_score": int(
            initial_settings.min_description_score
        ),
    }

    for key, default_value in chat_parameter_defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

    def reset_chat_options_to_page_settings() -> None:
        """Reset chatbot controls to the settings from the main page."""
        for key, default_value in chat_parameter_defaults.items():
            st.session_state[key] = default_value

    def clear_chat_history() -> None:
        """Clear previous chatbot questions and results."""
        st.session_state["product_chat_messages"] = []

    st.caption(
        "Enter a product description or part number to search the "
        "inventory immediately. For several products, enter one per line."
    )

    action_one, action_two, action_note = st.columns([1, 1.4, 2.6])

    with action_one:
        st.button(
            "Clear chat",
            key="clear_product_chat_dialog",
            on_click=clear_chat_history,
            use_container_width=True,
        )

    with action_two:
        st.button(
            "Use page settings",
            key="reset_chat_options",
            on_click=reset_chat_options_to_page_settings,
            use_container_width=True,
            help=(
                "Reset the chatbot filters to the supplier, stock, and "
                "maximum-result settings selected on the main page."
            ),
        )

    with action_note:
        st.caption(
            "The chatbot options apply only to new chatbot searches. "
            "They do not change the Excel-upload settings."
        )

    with st.expander("Search options", expanded=False):
        chat_search_all_suppliers = st.checkbox(
            "Search all suppliers",
            key="chat_search_all_suppliers",
            help=(
                "When selected, all suppliers are searched except those "
                "listed under Suppliers to exclude."
            ),
        )

        chat_preferred_suppliers = st.multiselect(
            "Preferred suppliers",
            options=available_suppliers,
            key="chat_preferred_suppliers",
            placeholder="Choose one or more suppliers",
            disabled=chat_search_all_suppliers,
            help=(
                "When Search all suppliers is not selected, only products "
                "from these suppliers are searched."
            ),
        )

        chat_excluded_suppliers = st.multiselect(
            "Suppliers to exclude",
            options=available_suppliers,
            key="chat_excluded_suppliers",
            placeholder="Optional — choose suppliers to remove",
            help=(
                "Excluded suppliers are removed after the preferred-supplier "
                "scope is applied. Exclusion takes priority."
            ),
        )

        stock_column, result_column = st.columns(2)

        with stock_column:
            chat_in_stock_only = st.checkbox(
                "Only products currently in stock",
                key="chat_in_stock_only",
                help=(
                    "Only inventory rows where OnHandQty is greater than "
                    "zero will be searched."
                ),
            )

        with result_column:
            chat_top_n = st.selectbox(
                "Maximum recommendations",
                options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
                key="chat_top_n",
                help=(
                    "Maximum number of recommendations returned for each "
                    "product description or part number."
                ),
            )

        chat_min_description_score = st.slider(
            "Minimum description match score",
            min_value=10,
            max_value=100,
            key="chat_min_description_score",
            help=(
                "The final description score must be at least the "
                "value selected here. This is the only minimum used for "
                "description matching. Fuzzy part-number matching separately "
                "uses its own threshold."
            ),
        )

    invalid_supplier_selection = (
        not chat_search_all_suppliers
        and not chat_preferred_suppliers
    )

    chat_settings = RecommendationSettings(
        preferred_suppliers=(
            ()
            if chat_search_all_suppliers
            else tuple(chat_preferred_suppliers)
        ),
        excluded_suppliers=tuple(chat_excluded_suppliers),
        in_stock_only=chat_in_stock_only,
        top_n=int(chat_top_n),
        min_description_score=float(
            chat_min_description_score
        ),
    )

    supplier_scope = (
        "All suppliers"
        if chat_search_all_suppliers
        else (
            f"{len(chat_preferred_suppliers)} preferred supplier"
            f"{'s' if len(chat_preferred_suppliers) != 1 else ''}"
        )
    )

    if chat_excluded_suppliers:
        supplier_scope += (
            f", {len(chat_excluded_suppliers)} excluded"
        )

    stock_scope = (
        "In-stock only"
        if chat_in_stock_only
        else "All stock levels"
    )

    st.info(
        f"Active chatbot filters — Suppliers: {supplier_scope} | "
        f"Stock: {stock_scope} | "
        f"Maximum results: {chat_top_n} | "
        f"Minimum description score: {chat_min_description_score}",
        icon="🔎",
    )

    if invalid_supplier_selection:
        st.warning(
            "Select at least one preferred supplier or enable "
            "Search all suppliers before searching."
        )

    # The result area appears above the chat input and scrolls independently.
    messages_container = st.container(height=430)

    chat_prompt = st.chat_input(
        "Enter a product description or part number",
        key="product_search_dialog_input",
        disabled=invalid_supplier_selection,
    )

    if chat_prompt:
        st.session_state["product_chat_messages"].append(
            {
                "role": "user",
                "content": chat_prompt,
            }
        )

        filter_summary = (
            f"Suppliers: {supplier_scope} | "
            f"Stock: {stock_scope} | "
            f"Maximum results: {chat_top_n} | "
            f"Minimum description score: {chat_min_description_score}"
        )

        try:
            with st.spinner(
                "Searching the inventory for the closest products..."
            ):
                (
                    chat_output_df,
                    chat_output_bytes,
                    chat_descriptions,
                ) = run_chat_product_search(
                    chat_prompt,
                    chat_settings,
                    inventory_part_numbers,
                )

            matched_chat_df = chat_output_df.copy()

            if "MatchType" in matched_chat_df.columns:
                matched_chat_df = matched_chat_df.loc[
                    matched_chat_df["MatchType"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .ne("")
                ].copy()

            if matched_chat_df.empty:
                assistant_text = (
                    "I could not find a credible recommendation under the "
                    "selected chatbot filters and minimum score."
                )
                displayed_results = pd.DataFrame()
            else:
                result_count = len(matched_chat_df)
                enquiry_count = len(chat_descriptions)

                assistant_text = (
                    f"I found {result_count} credible recommendation"
                    f"{'s' if result_count != 1 else ''} for "
                    f"{enquiry_count} search"
                    f"{'es' if enquiry_count != 1 else ''}."
                )

                chat_columns = get_chat_result_columns(
                    matched_chat_df
                )
                displayed_results = matched_chat_df[
                    chat_columns
                ].copy()

            st.session_state["product_chat_messages"].append(
                {
                    "role": "assistant",
                    "content": assistant_text,
                    "filter_summary": filter_summary,
                    "results": displayed_results,
                    "download_bytes": chat_output_bytes,
                    "download_name": (
                        "chat_product_recommendations.xlsx"
                    ),
                }
            )

        except Exception as exc:
            st.session_state["product_chat_messages"].append(
                {
                    "role": "assistant",
                    "content": (
                        "I could not complete the product search. "
                        f"Details: {exc}"
                    ),
                    "filter_summary": filter_summary,
                    "results": None,
                    "download_bytes": None,
                }
            )

    with messages_container:
        messages = st.session_state["product_chat_messages"]

        if not messages:
            with st.chat_message("assistant"):
                st.markdown(
                    "Enter a description such as **Sharpie red permanent "
                    "marker fine tip**, or enter a part number such as "
                    "**W2111A**. One product or part number per line."
                )

        for message_index, message in enumerate(messages):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

                if message.get("filter_summary"):
                    st.caption(
                        "Search settings: "
                        f"{message['filter_summary']}"
                    )

                message_results = message.get("results")

                if message_results is not None:
                    if message_results.empty:
                        st.info(
                            "Try adding a brand, model, size, colour, "
                            "pack quantity, or part number, or broaden the "
                            "supplier and stock filters."
                        )
                    else:
                        st.dataframe(
                            message_results,
                            use_container_width=True,
                            hide_index=True,
                        )

                download_bytes = message.get("download_bytes")

                if download_bytes is not None:
                    st.download_button(
                        "Download these recommendations",
                        data=download_bytes,
                        file_name=message["download_name"],
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key=(
                            "chat_dialog_download_"
                            f"{message_index}"
                        ),
                        use_container_width=True,
                    )


##st.title(APP_TITLE)

title_col, logo_col = st.columns([3, 2], vertical_alignment="center")

with title_col:
    st.title(APP_TITLE)

with logo_col:
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    st.image("assets/logo.gif", width="stretch")

st.caption(
    "Internal pilot — upload a customer enquiry, choose supplier scope, "
    "and download the recommended inventory matches."
)

##st.warning(
##    "Pilot only: this page has no user authentication. Run it only on an "
##   "approved internal network and stop the server when testing is finished.",
##    icon="⚠️",
##)

if not INVENTORY_FILE.exists():
    st.error(
        "Inventory file not found. Place the current inventory workbook at:\n\n"
        f"`{INVENTORY_FILE}`"
    )
    st.stop()

try:
    inventory_modified = INVENTORY_FILE.stat().st_mtime
    (
        available_suppliers,
        inventory_rows,
        inventory_part_numbers,
    ) = load_inventory_summary(
        str(INVENTORY_FILE),
        inventory_modified,
    )
except Exception as exc:
    st.error(f"Unable to read the inventory workbook: {exc}")
    st.stop()

inventory_timestamp = datetime.fromtimestamp(
    inventory_modified
).strftime("%d %b %Y, %I:%M %p")

metric_one, metric_two = st.columns(2)
metric_one.metric("Inventory rows", f"{inventory_rows:,}")
metric_two.metric("Inventory file updated", inventory_timestamp)

st.divider()

st.subheader("Customer enquiry")

enq_col1, enq_col2 = st.columns([2, 2], vertical_alignment="center")

with enq_col1:
    customer_template_bytes = generate_customer_enquiry_template()

    st.download_button(
        "Generate and download customer enquiry template",
        data=customer_template_bytes,
        file_name="customer_enquiry_template.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
        help=(
            "Generate a new Excel workbook containing the standard customer "
            "enquiry columns and instructions."
        ),
    )

with enq_col2:
    customer_file = st.file_uploader(
        "Customer enquiry workbook",
        type=["xlsx"],
        help=(
            "The workbook must contain a Description column. PartNumber, "
            "CustomerPrice, and PriceBasis are optional."
        ),
        on_change=clear_previous_results,
    )

st.divider()

all_suppliers = st.checkbox(
    "Search all suppliers",
    value=True,
    on_change=clear_previous_results,
)

selected_suppliers: list[str] = []
excluded_suppliers: list[str] = []

if not all_suppliers:

    sup_col1, sup_col2 = st.columns([2, 2], vertical_alignment="center")

    with sup_col1:
        selected_suppliers = st.multiselect(
            "Preferred suppliers",
            options=available_suppliers,
            placeholder="Choose one or more suppliers",
            help="Only products from these suppliers will be searched.",
            on_change=clear_previous_results,
        )

    with sup_col2:
        excluded_suppliers = st.multiselect(
            "Suppliers to exclude",
            options=available_suppliers,
            placeholder="Optional — choose suppliers to remove",
            help=(
                "Products from these suppliers will be removed from the search. "
                "Exclusion takes priority if a supplier is also selected as preferred."
            ),
            on_change=clear_previous_results,
        )

st.divider()

in_stock_only = st.checkbox(
    "Only recommend products currently in stock",
    value=True,
    help=(
        "When selected, only inventory rows with OnHandQty greater than zero "
        "are searched."
    ),
    on_change=clear_previous_results,
)

max_rec_col, min_scor_col = st.columns([2, 3], vertical_alignment="center")

with max_rec_col:
    top_n = st.selectbox(
        "Maximum recommendations per enquiry",
        options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        index=2,
        on_change=clear_previous_results,
    )

with min_scor_col:
    min_description_score = st.slider(
        "Minimum description match score",
        min_value=10,
        max_value=100,
        value=int(MIN_DESCRIPTION_SCORE),
        step=1,
        help=(
            "The final description score must be at least the value "
            "selected here. Higher is stricter and lower returns more possible "
            "matches. This is the only minimum used for description matching. "
            "Fuzzy part-number matching has a separate threshold."
        ),
        on_change=clear_previous_results,
    )

st.divider()


run_disabled = (
    customer_file is None
    or (not all_suppliers and not selected_suppliers)
)

current_supplier_scope = (
    "All suppliers"
    if all_suppliers
    else f"{len(selected_suppliers)} included"
)
if excluded_suppliers:
    current_supplier_scope += (
        f", {len(excluded_suppliers)} excluded"
    )

current_stock_scope = (
    "In-stock only"
    if in_stock_only
    else "All stock levels"
)

current_search_settings = RecommendationSettings(
    preferred_suppliers=(
        tuple(selected_suppliers)
        if not all_suppliers
        else ()
    ),
    excluded_suppliers=tuple(excluded_suppliers),
    in_stock_only=in_stock_only,
    top_n=int(top_n),
    min_description_score=float(
        min_description_score
    ),
)

generate_column, chat_column = st.columns([2, 1])

with generate_column:
    generate_clicked = st.button(
        "Generate recommendations",
        type="primary",
        disabled=run_disabled,
        use_container_width=True,
    )

with chat_column:
    chat_clicked = st.button(
        "💬 Product Search Bot",
        key="open_product_search_assistant",
        use_container_width=True,
        help=(
            "Open the Product Search Assistant. Its filters start "
            "with the current page settings and can then be changed inside "
            "the popup."
        ),
    )

if chat_clicked:
    st.session_state["chat_search_all_suppliers"] = (
        not bool(current_search_settings.preferred_suppliers)
    )
    st.session_state["chat_preferred_suppliers"] = list(
        current_search_settings.preferred_suppliers
    )
    st.session_state["chat_excluded_suppliers"] = list(
        current_search_settings.excluded_suppliers
    )
    st.session_state["chat_in_stock_only"] = (
        current_search_settings.in_stock_only
    )
    st.session_state["chat_top_n"] = (
        current_search_settings.top_n
    )
    st.session_state["chat_min_description_score"] = int(
        current_search_settings.min_description_score
    )

    show_product_search_assistant(
        current_search_settings,
        available_suppliers,
        inventory_part_numbers,
    )

if generate_clicked:
    clear_previous_results()

    try:
        with st.spinner("Matching customer enquiries to inventory..."):
            with TemporaryDirectory() as temp_directory:
                temp_path = Path(temp_directory)
                customer_path = temp_path / "customer_enquiry.xlsx"
                output_path = temp_path / "recommendations.xlsx"

                customer_path.write_bytes(customer_file.getvalue())

                settings = RecommendationSettings(
                    preferred_suppliers=(
                        tuple(selected_suppliers)
                        if not all_suppliers
                        else ()
                    ),
                    excluded_suppliers=tuple(excluded_suppliers),
                    in_stock_only=in_stock_only,
                    top_n=int(top_n),
                    min_description_score=float(
                        min_description_score
                    ),
                )

                output_df = run_recommendation(
                    customer_file=customer_path,
                    inventory_file=INVENTORY_FILE,
                    output_file=output_path,
                    settings=settings,
                )

                customer_df = pd.read_excel(customer_path)
                enquiry_count = (
                    customer_df["Description"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .ne("")
                    .sum()
                )

                st.session_state["recommendation_bytes"] = (
                    output_path.read_bytes()
                )
                st.session_state["recommendation_preview"] = output_df
                st.session_state["recommendation_filename"] = (
                    "recommendations.xlsx"
                )
                supplier_scope = (
                    "All suppliers"
                    if all_suppliers
                    else f"{len(selected_suppliers)} included"
                )
                if excluded_suppliers:
                    supplier_scope += f", {len(excluded_suppliers)} excluded"

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

                st.session_state["recommendation_metrics"] = {
                    "enquiries": int(enquiry_count),
                    "recommendations": int(
                        matched_recommendation_count
                    ),
                    "supplier_scope": supplier_scope,
                    "stock_scope": (
                        "In-stock only"
                        if in_stock_only
                        else "All stock levels"
                    ),
                    "minimum_score": int(
                        min_description_score
                    ),
                }

    except Exception as exc:
        st.error(f"Recommendation process failed: {exc}")

if "recommendation_bytes" in st.session_state:
    metrics = st.session_state["recommendation_metrics"]
    preview_df = st.session_state["recommendation_preview"]

    st.success("Recommendations generated successfully.")

    first, second, third, fourth = st.columns(4)
    first.metric("Enquiries processed", metrics["enquiries"])
    second.metric("Recommendations produced", metrics["recommendations"])
    third.metric("Supplier scope", metrics["supplier_scope"])
    fourth.metric("Stock scope", metrics["stock_scope"])

    st.caption(
        "Minimum description match score used for this run: "
        f"**{metrics['minimum_score']}**"
    )

    st.download_button(
        "Download recommendations.xlsx",
        data=st.session_state["recommendation_bytes"],
        file_name=st.session_state["recommendation_filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

    st.subheader("Result preview")

    if preview_df.empty:
        st.info(
            "No credible recommendations met the current matching rules."
        )
    else:
        preview_columns = [
            column
            for column in [
                "Customer_Description",
                "Rank",
                "Score",
                "Inventory_Product",
                "Inventory_Description",
                "Inventory_Supplier",
                "Inventory_Unit",
                "Inventory_ListPriceEx",
                "Inventory_OnHandQty",
                "MatchType"
            ]
            if column in preview_df.columns
        ]

        result_preview = preview_df[preview_columns]

        styled_preview = result_preview.style.set_properties(
            **{
                "background-color": "#FFFFFF",
                "color": "#333333",
            }
        )

        st.dataframe(
            styled_preview,
            use_container_width=True,
            hide_index=True,
        )

with st.expander("Pilot notes"):
    st.markdown(
        """
        - Download the standard customer enquiry template before entering customer requests.
        - The inventory workbook is controlled locally and is not uploaded by users.
        - Supplier names come directly from the inventory `Supplier` column.
        - Selecting no preferred supplier is not allowed unless **Search all suppliers** is enabled.
        - Excluded suppliers are removed after the preferred-supplier scope is applied.
        - **Only recommend products currently in stock** keeps rows where `OnHandQty > 0`.
        - The minimum description match score defaults to **80** and can be changed for both batch and chatbot searches.
        - The selected minimum score is the only threshold used for description matching.
        - Lowering the minimum makes lower-scoring description matches eligible; the highest-ranked eligible products are returned first.
        - Nonblank enquiries with no qualifying match remain in the Excel output with blank recommendation fields.
        - The Product Search Assistant has its own supplier, stock, and maximum-result controls. It uses the same deterministic local matching engine and does not call an external chatbot or AI API.
        - Closing the terminal that launched Streamlit stops the pilot server.
        """
    )