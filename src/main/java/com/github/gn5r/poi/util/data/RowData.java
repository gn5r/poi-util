package com.github.gn5r.poi.util.data;

import java.util.List;

import lombok.Data;

/**
 * 1行に関するデータオブジェクト
 *
 * @author gn5r
 */
@Data
public class RowData {

    /** 開始行インデックス */
    private Integer firstRow;

    /** 終了行インデックス */
    private Integer lastRow;

    /** 行の高さ */
    private Integer height;

    /** セルオブジェクト */
    private List<CellData> cells;

    /** フィールド名 */
    private String fieldName;
}
