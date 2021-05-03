package com.github.gn5r.poi.util.data;

import lombok.Data;

/**
 * 1セルに関するデータオブジェクト
 * 
 * @author gn5r
 */
@Data
public class CellData {

    /** 開始行インデックス */
    private Integer firstRow;

    /** 終了行インデックス */
    private Integer lastRow;

    /** 開始列インデックス */
    private int firstColl;

    /** 終了列インデックス */
    private int lastColl;

    /** 名前定義 */
    private String cellName;

    /** 値 */
    private Object value;
}
