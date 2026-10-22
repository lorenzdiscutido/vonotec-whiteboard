# Graph Report - explore_anti_gravity_features  (2026-10-22)

## Corpus Check
- Corpus is ~25,586 words - fits in a single context window. You may not need a graph.

## Summary
- 465 nodes · 989 edges · 24 communities (19 shown, 5 thin omitted)
- Extraction: 99% EXTRACTED · 1% INFERRED · 0% AMBIGUOUS · INFERRED: 13 edges (avg confidence: 0.85)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Batch Processing Pipeline
- SharePoint Graph Integration
- Dashboard Data Models
- Excel Report Generation
- AI Extraction & Gemini
- Defect Rules & Limits
- UI Base Components
- Main App Entrypoint
- UI Workflow Sections
- Dashboard Defect Analytics
- Concurrent File Loader
- SVG Icon System
- Source File Browser
- Filter & Header Controls
- File Cache & Scope
- Dimension Table Views
- Image & Size Formatting
- Substrate & Defect Filters
- Data Health & Review Banners
- Audit Log & Data Tables
- File Management Controls
- Result Row Rendering
- Action Footers & Buttons
- Status Summary Metrics

## God Nodes (most connected - your core abstractions)
1. `render_analytics_dashboard()` - 18 edges
2. `parse_workbook()` - 16 edges
3. `_get_access_token()` - 15 edges
4. `run_batch()` - 14 edges
5. `_get_site_id()` - 14 edges
6. `dimension_by_defect()` - 13 edges
7. `render_source_browser()` - 12 edges
8. `_run_and_finish()` - 11 edges
9. `_process_photo()` - 11 edges
10. `write_parsed_data_to_excel()` - 11 edges

## Surprising Connections (you probably didn't know these)
- `_prepare_output()` --calls--> `sanitize_batch_filename()`  [EXTRACTED]
  app.py → file_utils.py
- `_prepare_output()` --calls--> `resolve_subfolder()`  [EXTRACTED]
  app.py → ui_sections.py
- `_run_and_finish()` --calls--> `BatchOutcome`  [EXTRACTED]
  app.py → batch_processor.py
- `_run_and_finish()` --calls--> `run_batch()`  [EXTRACTED]
  app.py → batch_processor.py
- `_run_and_finish()` --calls--> `save_prefs()`  [EXTRACTED]
  app.py → file_utils.py

## Import Cycles
- None detected.

## Communities (24 total, 5 thin omitted)

### Community 0 - "Batch Processing Pipeline"
Cohesion: 0.06
Nodes (25): on_progress(), BatchOutcome, _create_folder(), _download_existing(), _process_photo(), _record_added(), _review_message(), run_batch() (+17 more)

### Community 1 - "SharePoint Graph Integration"
Cohesion: 0.08
Nodes (27): _children_url(), _create_folder_if_missing(), download_file(), download_file_to_bytes(), download_item_to_bytes(), _encode_folder_path(), _encode_item_path(), ensure_subfolder() (+19 more)

### Community 2 - "Dashboard Data Models"
Cohesion: 0.08
Nodes (20): _base_code(), _canonical_modifier(), cause_for_code(), _clean_token(), _convert_board_dimension(), dimension_by_defect(), _is_empty(), material_for_code() (+12 more)

### Community 3 - "Excel Report Generation"
Cohesion: 0.10
Nodes (14): _base_code_expr(), _build_rows_data(), _create_master_workbook(), _dynamic_repair_formula(), _expand_entries(), _layout(), _modifier_expr(), _populate_reference_sheet() (+6 more)

### Community 4 - "AI Extraction & Gemini"
Cohesion: 0.11
Nodes (18): extract_data(), _apply_unit(), _defect_reference_text(), _extraction_prompt(), find_problems(), _fix_dimension_units(), _generate_text(), _get_model() (+10 more)

### Community 5 - "Defect Rules & Limits"
Cohesion: 0.09
Nodes (9): _limit(), _build_defect_units(), calculate_repair(), _decimal(), _mround(), repair_rule(), _roundup(), _value() (+1 more)

### Community 6 - "UI Base Components"
Cohesion: 0.11
Nodes (13): _brand_html(), card_heading(), load_local_css(), render_empty_results(), render_header(), render_instructions(), render_stepper(), render_tip() (+5 more)

### Community 7 - "Main App Entrypoint"
Cohesion: 0.13
Nodes (9): _fail(), _prepare_output(), RemoteFile, _run_and_finish(), start_batch(), start_remote_batch(), hide_uploader_chrome_if_filled(), init_session_state() (+1 more)

### Community 8 - "UI Workflow Sections"
Cohesion: 0.15
Nodes (9): _direct_parent_label(), _get_nested_folders(), process_blockers(), _push_pending(), refresh_folder_list(), render_folder_picker(), render_pending_syncs(), resolve_subfolder() (+1 more)

### Community 9 - "Dashboard Defect Analytics"
Cohesion: 0.17
Nodes (8): count_by(), natural_key(), _bar_chart(), _chart_badge_text(), _chart_notes(), _chart_section(), _horizontal_bar_chart(), _sorted_options()

### Community 10 - "Concurrent File Loader"
Cohesion: 0.22
Nodes (5): _download_and_parse(), get_file_rows(), _load_one(), start_warmup_if_due(), _warm_up()

### Community 11 - "SVG Icon System"
Cohesion: 0.14
Nodes (5): icon(), configure_page(), empty_state(), live_chip(), load_polish_css()

### Community 12 - "Source File Browser"
Cohesion: 0.20
Nodes (8): _ago(), _cached_folders(), _footer_html_and_count(), _header_html(), _level_key(), render_source_browser(), _root(), _site_box_html()

### Community 13 - "Filter & Header Controls"
Cohesion: 0.21
Nodes (5): apply_filters(), section_heading(), _kpi_card_html(), render_analytics_dashboard(), _render_kpi_cards()

### Community 14 - "File Cache & Scope"
Cohesion: 0.20
Nodes (5): is_loaded(), _cached_file_list(), _cached_subfolders(), _folders_for_scope(), _load_scope()

### Community 15 - "Dimension Table Views"
Cohesion: 0.22
Nodes (5): _defect_description(), _dimension_section(), _dimension_table_html(), _number_text(), _summary_to_excel_bytes()

### Community 16 - "Image & Size Formatting"
Cohesion: 0.22
Nodes (4): _format_size(), render_photo_previews(), draw_grid(), _thumbnail_bytes()

### Community 17 - "Substrate & Defect Filters"
Cohesion: 0.22
Nodes (5): _defect_options_for_substrates(), _filter_key(), _refresh_lists(), render_dashboard_controls(), render_dashboard_view()

### Community 18 - "Data Health & Review Banners"
Cohesion: 0.29
Nodes (4): _base_defect_code(), _data_health_html(), _missing_dimension_rows(), _render_missing_dimensions_queue()

## Knowledge Gaps
- **5 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `dimension_by_defect()` connect `Dashboard Data Models` to `Dashboard Defect Analytics`, `Filter & Header Controls`?**
  _High betweenness centrality (0.019) - this node is a cross-community bridge._
- **Should `Batch Processing Pipeline` be split into smaller, more focused modules?**
  _Cohesion score 0.06246799795186892 - nodes in this community are weakly interconnected._
- **Why does `sanitize_batch_filename()` connect `Batch Processing Pipeline` to `Dashboard Data Models`, `Main App Entrypoint`?**
  _High betweenness centrality (0.019) - this node is a cross-community bridge._
- **Should `SharePoint Graph Integration` be split into smaller, more focused modules?**
  _Cohesion score 0.08417508417508418 - nodes in this community are weakly interconnected._
- **Why does `parse_workbook()` connect `Dashboard Data Models` to `Concurrent File Loader`, `Defect Rules & Limits`?**
  _High betweenness centrality (0.018) - this node is a cross-community bridge._
- **Should `Dashboard Data Models` be split into smaller, more focused modules?**
  _Cohesion score 0.07610993657505286 - nodes in this community are weakly interconnected._
- **Should `Excel Report Generation` be split into smaller, more focused modules?**
  _Cohesion score 0.09803921568627451 - nodes in this community are weakly interconnected._