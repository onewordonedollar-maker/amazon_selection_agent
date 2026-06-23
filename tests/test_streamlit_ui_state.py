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
        self.assertIn("render_cards(current_page_products, current_page_display_start)", source)
        self.assertIn("table_rows(current_page_products, current_page_display_start)", source)

    def test_result_exports_support_selected_current_page_and_all_filtered(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("result_export_scope", source)
        self.assertIn("export_products_for_scope", source)
        self.assertIn('"导出已勾选"', source)
        self.assertIn('"导出当前页"', source)
        self.assertIn('"导出全部结果"', source)

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
        self.assertIn('render_cards(current_page_products, current_page_display_start)', result_area)
        self.assertIn('render_tile_cards(current_page_products)', result_area)

    def test_tile_cards_show_rating_review_count_without_badges(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        tile_html = source[source.index("def product_tile_html"):source.index("def render_tile_cards")]

        self.assertIn("评分/评分数", tile_html)
        self.assertIn("rating_review_label", tile_html)
        self.assertNotIn("level-corner", tile_html)
        self.assertNotIn("tag tag-", tile_html)

    def test_tile_cards_use_inline_favorite_and_note_without_bottom_controls(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        tile_html = source[source.index("def product_tile_html"):source.index("def render_tile_cards")]
        tile_renderer = source[source.index("def render_tile_cards"):source.index("def render_favorites_panel")]
        favorites_renderer = source[source.index("def render_favorites_panel"):source.index("def collect_sellersprite_products")]
        tile_css = source[source.index(".tile-card {"):source.index("div[data-testid=\"stElementContainer\"]:has(.product-card-select-anchor)")]

        self.assertIn("def render_tile_product_favorite_button", source)
        self.assertIn("def render_tile_product_note_input", source)
        self.assertIn("render_tile_product_favorite_button(product, \"tile\")", tile_renderer)
        self.assertIn("render_tile_product_note_input(product, \"tile\")", tile_renderer)
        self.assertIn("render_tile_product_favorite_button(product, \"favorite\")", favorites_renderer)
        self.assertIn("render_tile_product_note_input(product, \"favorite\")", favorites_renderer)
        self.assertNotIn("render_product_annotation_controls(product, \"tile\")", tile_renderer)
        self.assertNotIn("render_product_annotation_controls(product, \"favorite\")", favorites_renderer)
        self.assertIn("tile-favorite-host", tile_html)
        self.assertIn("tile-favorite-proxy", tile_html)
        self.assertIn("tile-note-host", tile_html)
        self.assertIn("tile-note-proxy", tile_html)
        self.assertIn("data-favorite-key", tile_html)
        self.assertIn("data-note-key", tile_html)
        self.assertIn(".tile-favorite-host {", source)
        self.assertIn(".tile-note-proxy {", source)
        self.assertIn("-webkit-line-clamp: 1", source)
        self.assertIn("max-height: 72px", source)
        self.assertIn("overflow-y: auto", source)
        self.assertIn("min-height: 510px", tile_css)

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
        self.assertIn("list-select-host", source)
        self.assertIn("list-select-proxy", source)
        self.assertIn("clickRealControl('div.st-key-row_include_' + asin", source)
        self.assertIn(".seller-list-frame {", source)
        self.assertIn("padding: 0 0 6px;", source)

    def test_product_favorites_and_notes_are_wired_without_old_operation_icons(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        list_html = source[source.index("def seller_product_html"):source.index("def product_tile_html")]
        list_renderer = source[source.index("def render_cards"):source.index("def set_result_page")]
        tile_renderer = source[source.index("def render_tile_cards"):source.index("def collect_sellersprite_products")]

        self.assertIn("FAVORITES_PATH", source)
        self.assertIn("PRODUCT_NOTES_PATH", source)
        self.assertIn("def render_product_annotation_controls", source)
        self.assertIn("def render_list_product_favorite_button", source)
        self.assertIn("def render_list_product_note_input", source)
        self.assertIn("def render_tile_product_favorite_button", source)
        self.assertIn("def render_tile_product_note_input", source)
        self.assertIn("def render_list_favorite_portal", source)
        self.assertIn("toggle_product_favorite", source)
        self.assertIn("handle_product_note_change", source)
        self.assertIn("render_list_product_favorite_button(product)", list_renderer)
        self.assertIn("render_list_product_note_input(product)", list_renderer)
        self.assertNotIn("render_product_annotation_controls(product, \"list\")", list_renderer)
        self.assertIn("render_tile_product_favorite_button(product", tile_renderer)
        self.assertIn("render_tile_product_note_input(product", tile_renderer)
        self.assertNotIn("render_product_annotation_controls(product", tile_renderer)
        self.assertIn("tab_favorites", source)
        self.assertIn("render_favorites_panel", source)
        self.assertIn("list-note-host", list_html)
        self.assertIn("list-note-proxy", list_html)
        self.assertIn("data-note-value", list_html)
        self.assertIn("list-favorite-host", list_html)
        self.assertIn("product-list-favorite-anchor", source)
        self.assertIn("product-list-note-input-anchor", source)
        self.assertIn("render_list_favorite_portal()", source)
        self.assertNotIn("calc(100% - 54px)", source)
        self.assertNotIn("margin: -54px", source)
        self.assertNotIn("transform: translateY(94px)", source)
        self.assertIn("list-favorite-proxy", source)
        self.assertIn("list-note-proxy", source)
        self.assertIn("const favoriteKey = favoriteProxy.dataset.favoriteKey || ('favorite_list_' + asin)", source)
        self.assertIn("clickRealControl('div.st-key-' + favoriteKey + ' button')", source)
        self.assertIn("syncListNote", source)
        self.assertIn('if self.path == "/note"', source)
        self.assertIn("save_notes(PRODUCT_NOTES_PATH, notes)", source)
        self.assertIn("NOTE_SAVE_PORT = 8766", source)
        self.assertIn("ensure_note_save_server()", source)
        self.assertIn("fetch('http://127.0.0.1:__NOTE_SAVE_PORT__/note'", source)
        self.assertIn('"★" if is_favorite else "☆"', source)
        self.assertIn('favorite_label = "★" if favorite_class else "☆"', list_html)

    def test_list_favorite_uses_theme_color_and_aligns_with_operation_row(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        favorite_css = source[
            source.index(".list-favorite-host {"):
            source.index(
                "div[data-testid=\"stElementContainer\"]:has(.product-list-note-input-anchor)",
                source.index(".list-favorite-host {"),
            )
        ]
        ops_css = source[source.index(".ops-cell {"):source.index(".seller-detail {")]

        self.assertIn("align-items: center", favorite_css)
        self.assertIn("color: var(--muted-light) !important", favorite_css)
        self.assertIn("color: var(--brand) !important", favorite_css)
        self.assertIn("font-size: 24px !important", favorite_css)
        self.assertNotIn("#f2a900", favorite_css)
        self.assertIn("align-self: center", ops_css)

    def test_list_product_image_has_no_signal_badges_and_uses_larger_square(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        list_html = source[source.index("def seller_product_html"):source.index("def product_tile_html")]
        list_image_css = source[
            source.index(".seller-product {"):
            source.index(".seller-info {", source.index(".seller-product {"))
        ]

        self.assertNotIn("signal-tags", source)
        self.assertNotIn("tag-bs", source)
        self.assertNotIn("tag-ac", source)
        self.assertNotIn("tag-nr", source)
        self.assertNotIn(">BS<", list_html)
        self.assertNotIn(">AC<", list_html)
        self.assertNotIn(">NR<", list_html)
        self.assertIn("grid-template-columns: 124px minmax(0, 1fr)", list_image_css)
        self.assertIn("width: 116px", list_image_css)
        self.assertIn("height: 116px", list_image_css)
        self.assertIn("object-fit: contain", list_image_css)

    def test_list_note_defaults_to_plain_text_and_only_edits_on_demand(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        note_renderer = source[
            source.index("def render_list_product_note_input"):
            source.index("def render_product_annotation_controls")
        ]
        editor_css = source[
            source.index("div[data-testid=\"stElementContainer\"]:has(.product-list-note-input-anchor) + div[data-testid=\"stTextInput\"]"):
            source.index(".list-note-host {")
        ]
        editor_input_css = source[
            source.index("div[data-testid=\"stElementContainer\"]:has(.product-list-note-input-anchor) + div[data-testid=\"stTextInput\"] input"):
            source.index("div[data-testid=\"stElementContainer\"]:has(.product-annotation-anchor)")
        ]
        note_css = source[
            source.index(".list-note-host {"):
            source.index(
                "div[data-testid=\"stElementContainer\"]:has(.product-annotation-anchor)",
                source.index(".list-note-host {"),
            )
        ]

        self.assertIn("note_key = f\"product_note_list_{asin}\"", note_renderer)
        self.assertIn("product-list-note-input-anchor", note_renderer)
        self.assertIn("st.text_input(", note_renderer)
        self.assertIn("handle_product_note_change", note_renderer)
        self.assertIn("contenteditable", source)
        self.assertIn("blur", source)
        self.assertIn("input.dispatchEvent(new Event('input'", source)
        self.assertIn("input.dispatchEvent(new Event('change'", source)
        self.assertNotIn("st.form(", note_renderer)
        self.assertNotIn("st.form_submit_button(", note_renderer)
        self.assertNotIn("product_note_form_", source)
        self.assertNotIn("product_note_commit_", source)
        self.assertIn("position: absolute !important", editor_css)
        self.assertIn("opacity: 0 !important", editor_css)
        self.assertIn("width: 1px !important", editor_input_css)
        self.assertIn("height: 1px !important", editor_input_css)
        self.assertIn("background: transparent !important", note_css)
        self.assertIn("border: 0 !important", note_css)
        self.assertIn(".list-note-proxy", note_css)
        self.assertIn("text-align: left !important", note_css)
        self.assertIn("width: fit-content !important", note_css)
        self.assertNotIn("min-height: 44px", note_css)

    def test_result_cards_and_table_use_display_numbers_not_source_rank(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        list_renderer = source[source.index("def render_cards"):source.index("def set_result_page")]
        table_renderer = source[source.index("def table_rows"):source.index("def image_formula")]
        start = source.index("current_page_products = page_slice(products")
        end = source.index("if st.session_state.last_collection_summary", start)
        result_area = source[start:end]

        self.assertIn("page_start_index", source)
        self.assertIn("current_page_display_start", result_area)
        self.assertIn("render_cards(current_page_products, current_page_display_start)", source)
        self.assertIn("table_rows(current_page_products, current_page_display_start)", source)
        self.assertIn("display_number", list_renderer)
        self.assertNotIn('class="seller-rank">{product.rank}</div>', list_renderer)
        self.assertIn('row["rank"] = display_number', table_renderer)

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

    def test_list_product_header_stays_sticky_for_current_streamlit_container(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn('div[data-testid="stElementContainer"]:has(.seller-table-header-anchor) + div[data-testid="stElementContainer"]', source)
        direct_header_rule = source[
            source.index('div[data-testid="stElementContainer"]:has(.seller-table-header-anchor) + div[data-testid="stElementContainer"]'):
            source.index('div[data-testid="stPopover"] button', source.index('div[data-testid="stElementContainer"]:has(.seller-table-header-anchor) + div[data-testid="stElementContainer"]'))
        ]

        self.assertIn("position: sticky !important", direct_header_rule)
        self.assertIn("top: 80px !important", direct_header_rule)
        self.assertIn("z-index: 79 !important", direct_header_rule)

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

    def test_final_theme_keeps_stmain_overflow_visible_for_sticky_headers(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        final_theme = source[source.index("/* Visual refresh: presentation only. Core collection and filtering logic is unchanged. */"):]

        self.assertIn('body:not(:has(div[data-testid="stDialog"])) section[data-testid="stMain"]', final_theme)
        self.assertIn("overflow-y: visible !important", final_theme)
        self.assertIn("overflow-x: visible !important", final_theme)


if __name__ == "__main__":
    unittest.main()
