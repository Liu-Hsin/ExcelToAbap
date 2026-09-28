class ZCL_EXCEL2ABAP definition
  public
  final
  create private .

public section.

  class-methods UPLOAD_FILE
    importing
      !IM_FILE_NAME type STRING
      !IM_SHEET_NAME type STRING optional
    returning
      value(CT_TAB) type ref to DATA
    raising
      CX_ROOT .
  protected section.
  private section.

    types:
      begin of ty_xml_key_value,
        key   type string,  " 单元格坐标，如 A1
        row   type syindex, " 行号
        col   type syindex, " 列号
        value type string,  " 单元格值
      end of ty_xml_key_value .
    types:
      begin of ty_xml_node_name,
        type  type string,  " 'sheet' 或 'sharedString'
        index type syindex, " sheet: 序号(1起) / sharedString:si索引(0起)
        name  type string,  " sheet 名 / 字符串文本
        rid   type string,  " sheet 的 r:id，如 rId1
      end of ty_xml_node_name .
    types:
      begin of ty_style_fmt,
        fmt_type type string, " ''=普通数字 / 'D'=日期 / 'T'=时间 / 'DT'=日期时间
      end of ty_style_fmt .
    types:
      begin of ty_custom_fmt,
        numfmtid type string,
        fmt_type type string,
      end of ty_custom_fmt .
    types:
      begin of ty_merge,
        row_from type syindex,
        col_from type syindex,
        row_to   type syindex,
        col_to   type syindex,
      end of ty_merge .

    constants c_rels_ns type string value 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ##NO_TEXT.
    constants c_uri type string value 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ##NO_TEXT.
    constants c_sheet type string value 'sheet' ##NO_TEXT.
    constants c_sharedstring type string value 'sharedString' ##NO_TEXT.
    constants c_tagname_id type string value 'id' ##NO_TEXT.
    constants c_tagname_t type string value 't' ##NO_TEXT.
    constants c_tagname_r type string value 'r' ##NO_TEXT.
    constants c_tagname_c type string value 'c' ##NO_TEXT.
    constants c_tagname_f type string value 'f' ##NO_TEXT.
    constants c_tagname_v type string value 'v' ##NO_TEXT.
    constants c_tagname_s type string value 's' ##NO_TEXT.
    constants c_tagname_row type string value 'row' ##NO_TEXT.
    constants c_tagname_si type string value 'si' ##NO_TEXT.
    constants c_tagname_is type string value 'is' ##NO_TEXT.
    constants c_tagname_mergecell type string value 'mergeCell' ##NO_TEXT.
    constants c_tagname_ref type string value 'ref' ##NO_TEXT.
    constants c_tagname_numfmt type string value 'numFmt' ##NO_TEXT.
    constants c_tagname_numfmtid type string value 'numFmtId' ##NO_TEXT.
    constants c_tagname_formatcode type string value 'formatCode' ##NO_TEXT.
    constants c_tagname_cellxfs type string value 'cellXfs' ##NO_TEXT.
    constants c_item_name type string value 'name' ##NO_TEXT.
    constants c_mode_sheet type string value 'SHEET' ##NO_TEXT.
    constants c_mode_shared type string value 'SHARED' ##NO_TEXT.
    data:
      lt_xml_node_name type standard table of ty_xml_node_name with empty key .
    data:
      lt_xml_key_value type standard table of ty_xml_key_value with empty key .
    data:
      lt_style_fmt     type standard table of ty_style_fmt with empty key .
    data:
      lt_merges type standard table of ty_merge with empty key .
    data lo_document type ref to cl_xlsx_document .
    data mv_sheet_name type string .
    data mv_mode type string .
    data mv_si_index type syindex .
    data mv_cur_row type syindex .
    data mv_cur_cell type string .
    data mv_cur_type type string .
    data mv_cur_formula type string .
    data mv_cur_value type string .
    data mv_cur_style type string .

    methods constructor
      importing
        !im_file_name type string
      raising
        cx_root .
    methods opend_xml .
    methods worksheet
      importing
        !io_workbook type ref to cl_xlsx_workbookpart .
    methods parse_sheet_names
      importing
        !io_workbook type ref to cl_xlsx_workbookpart .
    methods shared_strings_part
      importing
        !io_workbook type ref to cl_xlsx_workbookpart .
    methods parse_styles
      importing
        !io_workbook type ref to cl_xlsx_workbookpart .
    methods read_sheet
      importing
        !io_part type ref to cl_openxml_part .
    methods parse_node
      importing
        !io_node type ref to if_ixml_node .
    methods create_xml_object
      importing
        !im_xml       type xstring
      returning
        value(ro_doc) type ref to if_ixml_document .
    methods get_attr_value
      importing
        !io_map       type ref to if_ixml_named_node_map
        !iv_name      type string
      returning
        value(rv_val) type string .
    methods resolve_cell_value .
    methods append_cell .
    methods append_shared_string .
    methods parse_cell_ref
      importing
        !iv_ref type string
      exporting
        !ev_row type syindex
        !ev_col type syindex .
    methods get_attr_value_ns
      importing
        !io_map       type ref to if_ixml_named_node_map
        !iv_uri       type string
        !iv_name      type string
      returning
        value(rv_val) type string .
    methods classify_fmt
      importing
        !iv_code       type string
      returning
        value(rv_type) type string .
    methods create_dynamic_tab
      returning
        value(rt_tab) type ref to data .
    methods get_col_letters
      importing
        !iv_key           type string
      returning
        value(rv_letters) type string .
    methods col_to_letter
      importing
        !iv_col           type syindex
      returning
        value(rv_letters) type string .

ENDCLASS.



CLASS ZCL_EXCEL2ABAP IMPLEMENTATION.


  method APPEND_CELL.

    data ls_kv type ty_xml_key_value.

    parse_cell_ref( exporting iv_ref = mv_cur_cell
                    importing ev_row = ls_kv-row
                              ev_col = ls_kv-col ).
    ls_kv-key   = mv_cur_cell.
    ls_kv-value = mv_cur_value.
    append ls_kv to lt_xml_key_value.

  endmethod.


  method APPEND_SHARED_STRING.

    data ls_node type ty_xml_node_name.

    ls_node-type  = c_sharedstring.
    ls_node-index = mv_si_index - 1.           " 第一个 si → 0
    ls_node-name  = mv_cur_value.
    append ls_node to lt_xml_node_name.

  endmethod.


  method CLASSIFY_FMT.

    data(lv_code) = to_upper( iv_code ).
    " 含年（yy）→ 至少是日期；含小时/秒（hh/ss）→ 至少是时间
    if lv_code cs 'YY' or lv_code cs 'Y'.
      rv_type = 'D'.
    endif.
    if lv_code cs 'HH' or lv_code cs 'SS' or lv_code cs 'H' or lv_code cs 'S'.
      if rv_type = 'D'.
        rv_type = 'DT'.
      else.
        rv_type = 'T'.
      endif.
    endif.

  endmethod.


  method COL_TO_LETTER.

    " 1→A, 26→Z, 27→AA
    data(lv_col) = iv_col.
    data lv_letters type string value 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
    while lv_col > 0.
      data(lv_rem) = ( lv_col - 1 ) mod 26.
      rv_letters = lv_letters+lv_rem(1) && rv_letters.
      lv_col = ( lv_col - 1 ) div 26.
    endwhile.

  endmethod.


  method CONSTRUCTOR.

    lo_document = cl_xlsx_document=>load_document(
      iv_data = cl_openxml_helper=>load_local_file( im_file_name = im_file_name ) ).

  endmethod.


  method CREATE_DYNAMIC_TAB.

    types:
      begin of ty_col_info,
        col     type syindex, " 列号，排序依据
        letters type string,  " 列字母，字段名
      end of ty_col_info.

    data lt_cols type standard table of ty_col_info with empty key.

    " 收集列：col + 字母
    loop at lt_xml_key_value assigning field-symbol(<fs_kv>).
      if not line_exists( lt_cols[ col = <fs_kv>-col ] ).
        data ls_col type ty_col_info.
        ls_col-col     = <fs_kv>-col.
        ls_col-letters = get_col_letters( <fs_kv>-key ). " 'A1' → 'A'
        append ls_col to lt_cols.
      endif.
    endloop.
    sort lt_cols by col.

    if lt_cols is initial.
      return.
    endif.

    " 动态结构：字段名直接用收集到的字母
    data lt_comp type abap_component_tab.
    data ls_comp type abap_componentdescr.
    loop at lt_cols into data(ls_col2).
      ls_comp-name = ls_col2-letters.
      ls_comp-type = cl_abap_elemdescr=>get_string( ).
      append ls_comp to lt_comp.
    endloop.

    data(lo_struct) = cl_abap_structdescr=>create( lt_comp ).
    data(lo_tab) = cl_abap_tabledescr=>create( p_line_type  = lo_struct
                                               p_table_kind = cl_abap_tabledescr=>tablekind_std ).
    create data rt_tab type handle lo_tab.


    data(lv_max_row) = 0.
    loop at lt_xml_key_value assigning <fs_kv>.
      lv_max_row = nmax( val1 = lv_max_row
                         val2 = <fs_kv>-row ).
    endloop.

    field-symbols <lt_data> type standard table.

    assign rt_tab->* to <lt_data>.
    do lv_max_row times.
      append initial line to <lt_data> assigning field-symbol(<fs_row>).
    enddo.

    "填值：字段名 = key 的字母前缀
    loop at lt_xml_key_value assigning <fs_kv>.
      assign <lt_data>[ <fs_kv>-row ] to <fs_row>.
      if sy-subrc = 0.
        assign component get_col_letters( <fs_kv>-key ) of structure <fs_row> to field-symbol(<lv_val>).
        if sy-subrc = 0.
          <lv_val> = <fs_kv>-value.
        endif.
      endif.
    endloop.

    " 展开合并单元格
    loop at lt_merges assigning field-symbol(<fs_m>).
      " 主单元格（左上角）的值
      read table <lt_data> assigning field-symbol(<fs_master_row>) index <fs_m>-row_from.
      if sy-subrc <> 0.
        continue.
      endif.
      assign component col_to_letter( <fs_m>-col_from )
             of structure <fs_master_row> to field-symbol(<fs_master_val>).
      if sy-subrc <> 0 or <fs_master_val> is initial.
        continue. " 主单元格本身没值，跳过
      endif.

      do <fs_m>-row_to - <fs_m>-row_from + 1 times.
        data(lv_r) = <fs_m>-row_from + sy-index - 1.
        read table <lt_data> assigning field-symbol(<fs_target_row>) index lv_r.
        if sy-subrc <> 0.
          continue.
        endif.
        do <fs_m>-col_to - <fs_m>-col_from + 1 times.
          data(lv_c) = <fs_m>-col_from + sy-index - 1.
          if lv_r = <fs_m>-row_from and lv_c = <fs_m>-col_from.
            continue." 主单元格本身，跳过
          endif.
          assign component col_to_letter( lv_c )
                 of structure <fs_target_row> to field-symbol(<fs_target_val>).
          if sy-subrc = 0.
            <fs_target_val> = <fs_master_val>.
          endif.
        enddo.
      enddo.
    endloop.

    "去除空行
    delete <lt_data> where ('table_line is initial').

  endmethod.


  method CREATE_XML_OBJECT.

    data(lo_ixml) = cl_ixml=>create( ).
    ro_doc = lo_ixml->create_document( ).
    data(lo_factory) = lo_ixml->create_stream_factory( ).
    data(lo_istream) = lo_factory->create_istream_xstring( im_xml ).

    lo_ixml->create_parser( document       = ro_doc
                            istream        = lo_istream
                            stream_factory = lo_factory )->parse( ).

  endmethod.


  method GET_ATTR_VALUE.

    "  属性可能不存在，统一保护
    if io_map->get_named_item( iv_name ) is not initial.
      rv_val = io_map->get_named_item( iv_name )->get_value( ).
    endif.

  endmethod.


  method GET_ATTR_VALUE_NS.

    if io_map->get_named_item_ns( uri  = iv_uri
                                  name = iv_name ) is not initial.
      rv_val = io_map->get_named_item_ns( uri  = iv_uri
                                          name = iv_name )->get_value( ).
    endif.

  endmethod.


  method GET_COL_LETTERS.

    " 'A1' → 'A'，'AB12' → 'AB'
    data(lv_rest) = iv_key.
    while lv_rest is not initial and lv_rest(1) between 'A' and 'Z'.
      rv_letters = rv_letters && lv_rest(1).
      lv_rest = lv_rest+1.
    endwhile.

  endmethod.


  method OPEND_XML.

    data(lo_workbook) = lo_document->get_workbookpart( ).
    "获取页签名称
    parse_sheet_names( lo_workbook ).
    "处理字符串
    shared_strings_part( lo_workbook ).
    "处理单元格样式
    parse_styles( lo_workbook ).
    "处理页签
    worksheet( lo_workbook ).

  endmethod.


  method PARSE_CELL_REF.

    " A1=>1,1 B1=>2,1 .. AA20=>27,20
    data(lv_rest) = iv_ref.
    data lv_letters type sy-abcde value 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.
    while lv_rest is not initial and lv_rest(1) between 'A' and 'Z'.
      data(lv_pos) = find( val = lv_letters
                           sub = lv_rest(1) ) + 1.
      ev_col = ev_col * 26 + lv_pos.
      lv_rest = lv_rest+1.
    endwhile.
    ev_row = lv_rest.
    " 剩余数字部分

  endmethod.


  method PARSE_NODE.

    data(lo_children) = io_node->get_children( ).

    do lo_children->get_length( ) times.
      data(lo_child) = lo_children->get_item( sy-index - 1 ).

      if lo_child->get_type( ) <> if_ixml_node=>co_node_element.
        " 跳过文本/空白节点
        continue.
      endif.
      data(lv_name) = lo_child->get_name( ).

      case lv_name.
        when c_tagname_row.
          mv_cur_row = get_attr_value( io_map  = lo_child->get_attributes( )
                                       iv_name = c_tagname_r ).
          parse_node( lo_child ).
        when c_tagname_c.
          "  聚合边界：进入清空 → 递归填充 → 退出落库
          if mv_mode <> c_mode_sheet.
            continue.
          endif.
          mv_cur_cell = to_upper( get_attr_value( io_map  = lo_child->get_attributes( )
                                                  iv_name = c_tagname_r ) ).
          mv_cur_type = get_attr_value( io_map  = lo_child->get_attributes( )
                                        iv_name = c_tagname_t ).
          mv_cur_style = get_attr_value( io_map  = lo_child->get_attributes( )
                                         iv_name = c_tagname_s ).
          clear: mv_cur_formula,
                 mv_cur_value.
          parse_node( lo_child ).
          resolve_cell_value( ).
          append_cell( ).
        when c_tagname_f.
          " 叶子：只取公式
          mv_cur_formula = lo_child->get_value( ).
        when c_tagname_v.
          " 叶子：只取值，不再拼接
          mv_cur_value = lo_child->get_value( ).

        when c_tagname_mergecell.
          " 合并单元格
          if mv_mode = c_mode_sheet.
            data(lv_ref) = get_attr_value( io_map = lo_child->get_attributes( ) iv_name = c_tagname_ref ).
            split lv_ref at ':' into data(lv_left) data(lv_right).
            parse_cell_ref( exporting iv_ref = lv_left
                            importing ev_row = data(lv_row1)
                                      ev_col = data(lv_col1) ).

            parse_cell_ref( exporting iv_ref = lv_right
                            importing ev_row = data(lv_row2)
                                      ev_col = data(lv_cow2) ).

            append value ty_merge( row_from = lv_row1 col_from = lv_col1
                                   row_to   = lv_row2 col_to   = lv_cow2 ) to lt_merges.
            clear:lv_ref,lv_left,lv_right,lv_row1,lv_col1,lv_row2,lv_cow2.
          endif.
        when c_tagname_is.
          " 内联字符串下钻
          parse_node( lo_child ).
        when c_tagname_si.
          "  聚合边界：按 si 计数
          if mv_mode <> c_mode_shared.
            continue.
          endif.
          clear mv_cur_value.
          mv_si_index += 1.
          parse_node( lo_child ).
          append_shared_string( ).
        when c_tagname_r.
          " 富文本 run：下钻找 t
          parse_node( lo_child ).
        when c_tagname_t.
          " 累加，兼容多 run
          mv_cur_value = mv_cur_value && lo_child->get_value( ).
        when others.
          " 中间层/未知标签：继续递归
          parse_node( lo_child ).
      endcase.
    enddo.

  endmethod.


  method PARSE_SHEET_NAMES.

    "workbook.xml：记录sheet 名
    data ls_node type ty_xml_node_name.

    data(l_xml) = io_workbook->get_data( ).
    data(lo_doc) = create_xml_object( l_xml ).
    data(lo_elements) = lo_doc->get_elements_by_tag_name_ns( name = c_sheet
                                                             uri  = c_uri ).
    data(lv_count) = lo_elements->get_length( ).
    do lv_count times.
      data(lo_node) = lo_elements->get_item( sy-index - 1 ).
      ls_node-type       = c_sheet.
      ls_node-index      = sy-index. " sheet 序号（1 起）
      ls_node-name       = to_upper( get_attr_value( io_map  = lo_node->get_attributes( )
                                                     iv_name = c_item_name ) ).
      ls_node-rid = get_attr_value_ns( io_map  = lo_node->get_attributes( )
                                       iv_uri  = c_rels_ns
                                       iv_name = c_tagname_id ).
      append ls_node to lt_xml_node_name.
      clear:lo_node, ls_node.
    enddo.

  endmethod.


  method PARSE_STYLES.

    data lv_data_type type string.

    data lt_custom    type standard table of ty_custom_fmt with empty key.
    data ls_custom    type ty_custom_fmt.

    data(l_xml)  = io_workbook->get_stylespart( )->get_data( ).
    data(lo_doc) = create_xml_object( l_xml ).
    data(lo_numfmts) = lo_doc->get_elements_by_tag_name_ns( name = c_tagname_numfmt
                                                            uri  = c_uri ).

    do lo_numfmts->get_length( ) times.
      data(lo_numfmt) = lo_numfmts->get_item( sy-index - 1 ).
      data(lo_attrs)  = lo_numfmt->get_attributes( ).

      ls_custom-numfmtid = get_attr_value( io_map  = lo_attrs
                                           iv_name = c_tagname_numfmtid ).
      ls_custom-fmt_type = classify_fmt( get_attr_value( io_map  = lo_attrs
                                                         iv_name = c_tagname_formatcode ) ).
      append ls_custom to lt_custom.
      clear: lo_attrs,
             lo_numfmt.
    enddo.

    data(lo_cellxfs_list) = lo_doc->get_elements_by_tag_name_ns( name = c_tagname_cellxfs
                                                                 uri  = c_uri ).
    if lo_cellxfs_list->get_length( ) = 0.
      return.
    endif.

    data(lo_cellxfs) = lo_cellxfs_list->get_item( 0 ).
    data(lo_xfs) = lo_cellxfs->get_children( ).

    do lo_xfs->get_length( ) times.
      data(lo_xf) = lo_xfs->get_item( sy-index - 1 ).
      if lo_xf->get_type( ) <> if_ixml_node=>co_node_element.
        continue.
      endif.

      data(lv_numfmtid) = get_attr_value( io_map  = lo_xf->get_attributes( )
                                          iv_name = c_tagname_numfmtid ).
      lv_data_type = ''.

      case lv_numfmtid. " 内置日期/时间格式 id（OOXML 标准）
        when '14' or '15' or '16' or '17' or '22' or '27' or '30' or '31' or '36' or '50' or '52' or '57'. " ....
          lv_data_type = 'D'.
        when '18' or '19' or '20' or '21' or '45' or '46' or '47'.
          lv_data_type = 'T'.
        when others.
          assign lt_custom[ numfmtid = lv_numfmtid ] to field-symbol(<fs_custom>).
          if sy-subrc = 0.
            lv_data_type = <fs_custom>-fmt_type.
          endif.
          unassign <fs_custom>.
      endcase.
      append value ty_style_fmt( fmt_type = lv_data_type ) to lt_style_fmt.
      clear: lv_data_type,
             lv_numfmtid,
             lo_xf.
    enddo.

  endmethod.


  method READ_SHEET.

    data(l_xml)  = io_part->get_data( ).
    data(lo_doc) = create_xml_object( l_xml ).
    mv_mode = c_mode_sheet.
    parse_node( lo_doc ).

  endmethod.


  method RESOLVE_CELL_VALUE.

    data: lv_date       type sydatum,
          lv_start_date type sydatum value '18991230',
          lv_time       type syuzeit,
          lv_start_time type syuzeit value '000000',
          lv_dec        type decfloat34,
          lv_int        type i,
          lv_frac       type decfloat34,
          lv_sec        type i.

    case mv_cur_type.
      when 's'.  " 共享字符串：索引 → 查表
        if mv_cur_value is not initial.
          assign lt_xml_node_name[ type  = c_sharedstring
                                   index = mv_cur_value ] to field-symbol(<fs_node>).
          if sy-subrc = 0.
            mv_cur_value = <fs_node>-name.
          endif.
        endif.
      when 'e'.    " 错误值 #N/A 等 → 标记
        mv_cur_value = |【错误】{ mv_cur_value }|.
      when 'inlineStr'. " 已是 <is><t> 的文本
      when others. " 数字/布尔/空：原样
    endcase.
    if mv_cur_type is not initial or mv_cur_style is initial.
      return.
    endif.

    assign lt_style_fmt[ mv_cur_style + 1 ] to field-symbol(<fs_fmt>)." s 从 0 起，表索引从 1 起
    if <fs_fmt> is not assigned.
      return.
    endif.
    if <fs_fmt>-fmt_type is initial.
      return." 普通数字，不动
    else.
      lv_dec = conv decfloat34( mv_cur_value ).
      lv_int = trunc( lv_dec ).
      case <fs_fmt>-fmt_type.
        when 'D'." dats 类型
          if lv_int is not initial.
            lv_date = lv_start_date + lv_int.
            mv_cur_value = lv_date.
          endif.
        when 'T'." time 类型
          lv_frac = lv_dec - lv_int.
          lv_sec = round( val = lv_frac * 86400
                          dec = 0 ).
          lv_time = lv_start_time + lv_sec.
          mv_cur_value = lv_time.
        when 'DT'.
          if lv_int is not initial.
            lv_date = lv_start_date + lv_int.
          endif.
          lv_frac = lv_dec - lv_int.
          lv_sec = round( val = lv_frac * 86400
                          dec = 0 ).
          lv_time = lv_start_time + lv_sec.
          mv_cur_value = |{ lv_date } { lv_time }|."20260101 165601
      endcase.
    endif.

  endmethod.


  method SHARED_STRINGS_PART.

    " sharedStrings.xml 记录单元格字符串
    data(l_xml) = io_workbook->get_sharedstringspart( )->get_data( ).
    data(lo_doc) = create_xml_object( l_xml ).
    mv_mode = c_mode_shared.
    mv_si_index = 0.
    parse_node( lo_doc ).

  endmethod.


  method UPLOAD_FILE.

    data(lo_excel) = new zcl_excel2abap( im_file_name ).

    if im_sheet_name is not initial.
      lo_excel->mv_sheet_name = to_upper( im_sheet_name ).
    endif.

    lo_excel->opend_xml( ).

    ct_tab = lo_excel->create_dynamic_tab( ).
    clear lo_excel.

  endmethod.


  method WORKSHEET.

    data(lo_worksheet) = io_workbook->get_worksheetparts( ).
    data(lv_count)     = lo_worksheet->get_count( ).
    check lv_count gt 0.

    if mv_sheet_name is not initial.
      assign lt_xml_node_name[ type = c_sheet
                               name = mv_sheet_name ] to field-symbol(<fs_node>).
    else.
      assign lt_xml_node_name[ type = c_sheet
                              index = 1 ] to <fs_node>.
    endif.
    if <fs_node> is assigned.
      read_sheet( io_workbook->get_part_by_id( <fs_node>-rid ) ).
    endif.
    " 指定了不存在的 sheet 名：循环结束，不读任何 sheet（lt_xml_key_value 为空）

  endmethod.
ENDCLASS.
