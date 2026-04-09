package org.jxls.transform.poi;

import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.common.SheetData;

/**
 * A streaming transformer that only loads sheets whose name ends with {@code _data} into the
 * transformation engine.  All other sheets in the workbook are left untouched and written to the
 * output as static (pass-through) sheets.
 *
 * <p>Usage pattern:
 * <ol>
 *   <li>Name each template sheet that contains jxls directives with the {@code _data} suffix
 *       (e.g. {@code Orders_data}).</li>
 *   <li>Leave every other sheet name as-is – those sheets will appear unchanged in the output.</li>
 *   <li>Use this transformer instead of {@link SelectSheetsForStreamingPoiTransformer}.</li>
 * </ol>
 */
public class DataSuffixStreamingPoiTransformer extends SelectSheetsForStreamingPoiTransformer {

    /** Suffix used to identify template sheets that should be processed by jxls. */
    public static final String DATA_SHEET_SUFFIX = "_data";

    public DataSuffixStreamingPoiTransformer(Workbook workbook) {
        super(workbook);
    }

    /**
     * @param allSheets whether <em>all</em> {@code _data} sheets should use SXSSF streaming writes;
     *                  note this controls the streaming write mode, not which sheets are loaded –
     *                  loading is always filtered to sheets ending with {@value #DATA_SHEET_SUFFIX}
     */
    public DataSuffixStreamingPoiTransformer(Workbook workbook, boolean allSheets,
            int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        super(workbook, allSheets, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
    }

    /**
     * @param sheetNames names of the {@code _data} sheets that should use SXSSF streaming writes;
     *                   note this controls the streaming write mode, not which sheets are loaded –
     *                   loading is always filtered to sheets ending with {@value #DATA_SHEET_SUFFIX}
     */
    public DataSuffixStreamingPoiTransformer(Workbook workbook, Set<String> sheetNames,
            int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        super(workbook, sheetNames, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
    }

    /**
     * Loads only sheets whose name ends with {@value #DATA_SHEET_SUFFIX} into the sheet map.
     * Other sheets remain in the workbook and are written to the output file unchanged.
     */
    @Override
    protected void readCellData() {
        Workbook wb = getWorkbook();
        int numberOfSheets = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet.getSheetName().endsWith(DATA_SHEET_SUFFIX)) {
                SheetData sheetData = PoiSheetData.createSheetData(sheet, this);
                sheetMap.put(sheetData.getSheetName(), sheetData);
            }
        }
    }
}
