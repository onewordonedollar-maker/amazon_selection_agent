from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class StreamlitUiStateTests(unittest.TestCase):
    def test_history_load_buttons_close_category_dialog_before_rerun(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("def close_category_dialog_state", source)
        self.assertIn('key="load_history_record_button"', source)
        self.assertIn('key="load_last_raw_button"', source)
        self.assertGreaterEqual(source.count("on_click=close_category_dialog_state"), 2)

    def test_category_row_selection_uses_compact_state_model(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("toggle_compact_category_selection", source)
        self.assertNotIn("def set_category_branch_selected", source)
        self.assertNotIn("def sync_category_select_all_state", source)
        self.assertNotIn("def sync_category_ancestors", source)

    def test_result_pagination_state_and_slicing_are_wired(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("PAGE_SIZE_OPTIONS", source)
        self.assertIn("result_page_size", source)
        self.assertIn("result_current_page", source)
        self.assertIn("current_page_products = page_slice(products", source)
        self.assertIn("render_cards(current_page_products)", source)
        self.assertIn("table_rows(current_page_products)", source)

    def test_result_exports_support_selected_current_page_and_all_filtered(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("result_export_scope", source)
        self.assertIn("export_products_for_scope", source)
        self.assertIn('"已勾选"', source)
        self.assertIn('"当前页"', source)
        self.assertIn('"全部筛选结果"', source)

    def test_bulk_select_all_results_does_not_write_every_product_widget_key(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("def select_all_filtered_products", source)
        self.assertIn("clear_product_selection_widget_state(products)", source)
        self.assertIn("visible_products", source)
        self.assertIn('st.session_state[f"row_include_{product.asin}"] = True', source)
        self.assertNotIn("set_all_product_selection(products, True)", source)

    def test_export_generation_is_cached_without_selection_state(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("def export_rows_signature", source)
        self.assertIn("@st.cache_data(show_spinner=False)", source)
        self.assertIn("def cached_excel_bytes", source)
        self.assertIn("def cached_csv_bytes", source)
        self.assertIn("row = {field: product_row.get(field, \"\") for field, _ in SELLERSPRITE_EXPORT_COLUMNS}", source)
        self.assertNotIn("asdict(product)", source[source.index("def export_rows_signature"):source.index("def cached_csv_bytes")])

    def test_select_all_results_does_not_force_a_second_rerun(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        select_all_block = source[
            source.index("actions_toolbar[2].button("):
            source.index('if actions_toolbar[3].button("复制ASIN"')
        ]

        self.assertIn("args=(products, current_page_products)", select_all_block)
        self.assertIn("on_click=select_all_filtered_products", select_all_block)
        self.assertNotIn("st.rerun()", select_all_block)

    def test_result_exports_are_lazy_not_generated_during_render(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        result_area = source[source.index("with tab_cards:"):source.index("with tab_log:")]

        self.assertIn("render_lazy_export_button", source)
        self.assertNotIn("data=excel_bytes(products)", result_area)
        self.assertNotIn("data=csv_bytes(products)", result_area)
        self.assertNotIn("data=excel_bytes(export_products)", result_area)
        self.assertNotIn("data=csv_bytes(export_products)", result_area)

    def test_product_results_support_list_and_tile_modes(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        result_area = source[source.index("with tab_cards:"):source.index("with tab_table:")]

        self.assertIn("result_view_mode", source)
        self.assertIn('["列表", "平铺"]', source)
        self.assertIn("def product_tile_html", source)
        self.assertIn("def render_tile_cards", source)
        self.assertIn('render_cards(current_page_products)', result_area)
        self.assertIn('render_tile_cards(current_page_products)', result_area)

    def test_tile_cards_show_rating_review_count_without_badges(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        tile_html = source[source.index("def product_tile_html"):source.index("def render_tile_cards")]

        self.assertIn("评分/评分数", tile_html)
        self.assertIn("rating_review_label", tile_html)
        self.assertNotIn("level-corner", tile_html)
        self.assertNotIn("tag tag-", tile_html)

    def test_product_selection_checkbox_is_inside_list_and_tile_cards(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        list_renderer = source[source.index("def render_cards"):source.index("def set_result_page")]
        tile_renderer = source[source.index("def render_tile_cards"):source.index("def collect_sellersprite_products")]

        self.assertIn("product-card-select-anchor", list_renderer)
        self.assertIn("tile-card-select-anchor", tile_renderer)
        self.assertNotIn("seller-select-cell", list_renderer)
        self.assertNotIn("row_select, row_body", list_renderer)
        self.assertIn("div[data-testid=\"stElementContainer\"]:has(.product-card-select-anchor) + div[data-testid=\"stCheckbox\"]", source)
        self.assertIn("div[data-testid=\"stElementContainer\"]:has(.tile-card-select-anchor) + div[data-testid=\"stCheckbox\"]", source)
        self.assertIn(".seller-list-frame {", source)
        self.assertIn("padding: 0 0 6px;", source)

    def test_result_view_mode_uses_compact_segmented_buttons(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        result_area = source[source.index("with tab_cards:"):source.index("with tab_table:")]

        self.assertIn("def set_result_view_mode", source)
        self.assertIn("result-view-toggle-anchor", source)
        self.assertIn('key="result_view_list_button"', result_area)
        self.assertIn('key="result_view_tile_button"', result_area)
        self.assertNotIn(".radio(", result_area)

    def test_result_toolbar_is_split_into_responsive_groups(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        result_area = source[source.index("with tab_cards:"):source.index("with tab_table:")]

        self.assertIn("cards-toolbar-actions-anchor", result_area)
        self.assertIn("cards-toolbar-controls-anchor", result_area)
        self.assertIn("actions_toolbar = st.columns", result_area)
        self.assertIn("controls_toolbar = st.columns", result_area)
        self.assertNotIn("\n        toolbar = st.columns", result_area)

    def test_result_toolbar_uses_compact_action_and_sort_labels(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        result_area = source[source.index("with tab_cards:"):source.index("with tab_table:")]

        self.assertIn("导出已勾选", source)
        self.assertIn("导出当前页", source)
        self.assertIn("导出全部结果", source)
        self.assertIn('"应用排序"', result_area)
        self.assertIn("cards-toolbar-sort-anchor", result_area)
        self.assertNotIn('"确定"', result_area)

    def test_lazy_export_button_does_not_stack_prepare_and_download(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        helper = source[source.index("def render_lazy_export_button"):source.index("def level_badge")]

        self.assertIn("download_label", helper)
        self.assertIn("else:", helper)
        self.assertLess(helper.index("container.download_button"), helper.index("else:"))

    def test_general_controls_close_stale_category_dialog_before_rerun(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("def close_category_dialog_state", source)
        self.assertIn('@st.dialog("选择类目", width="large", on_dismiss=close_category_dialog_state)', source)
        self.assertIn('key="ui_language_selector"', source)
        self.assertIn('key="list_type_selector"', source)
        self.assertIn('key="marketplace_selector"', source)
        self.assertGreaterEqual(source.count("on_change=close_category_dialog_state"), 3)

    def test_category_dialog_footer_has_scoped_layout_and_scroll_contract(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("category-dialog-grid-anchor", source)
        self.assertIn("category-footer-anchor", source)
        self.assertIn("max-height: max(220px, calc(100dvh - 250px))", source)
        self.assertIn("overflow-y: auto !important", source)
        self.assertNotIn('div[data-testid="stDialog"] div[data-testid="stHorizontalBlock"] {', source)

    def test_page_scroll_is_not_locked_when_category_dialog_is_closed(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn('body:not(:has(div[data-testid="stDialog"])) [data-testid="stAppViewContainer"]', source)
        self.assertIn('body:not(:has(div[data-testid="stDialog"])) section[data-testid="stMain"]', source)
        self.assertIn("overflow-y: visible !important", source)
        self.assertIn('body:has(div[data-testid="stDialog"]) [data-testid="stAppViewContainer"]', source)


if __name__ == "__main__":
    unittest.main()
