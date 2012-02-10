package com.jxls.writer;

import com.jxls.writer.command.Context;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 2/3/12 12:18 PM
 */
public class CellData {
    public static final String USER_FORMULA_PREFIX = "$[";
    public static final String USER_FORMULA_SUFFIX = "]";
    protected static final String REGEX_EXPRESSION = "\\$\\{[^}]*}";
    protected static final Pattern REGEX_EXPRESSION_PATTERN = Pattern.compile(REGEX_EXPRESSION);

    public enum CellType {
        STRING, NUMBER, BOOLEAN, DATE, FORMULA, BLANK, ERROR
    }

    protected String formula;
    protected int col;
    protected int row;
    protected String sheetName;
    protected Object evaluationResult;
    protected Object cellValue;

    protected CellType cellType;
    protected CellType targetCellType;

    List<Pos> targetPos = new ArrayList<Pos> ();



    public CellData(Pos pos) {
        this.sheetName = pos.getSheetName();
        this.row = pos.getRow();
        this.col = pos.getCol();
    }

    public CellData(String sheetName, int row, int col, CellType cellType, Object cellValue) {
        this.sheetName = sheetName;
        this.row = row;
        this.col = col;
        this.cellType = cellType;
        this.cellValue = cellValue;
        updateFormulaValue();
    }

    public CellData(Pos pos, CellType cellType, Object cellValue) {
        this.sheetName = pos.getSheetName();
        this.row = pos.getRow();
        this.col = pos.getCol();
        this.cellType = cellType;
        this.cellValue = cellValue;
        updateFormulaValue();
    }

    public CellData(String sheetName, int row, int col) {
        this(sheetName, row, col, CellType.BLANK, null);
    }


    public Object evaluate(Context context){
        targetCellType = cellType;
        if( cellType == CellType.STRING && cellValue != null){
            String strValue = cellValue.toString();
            if( isUserFormula(strValue) ){
                String formulaStr = strValue.substring(2, strValue.length()-1);
                evaluate(formulaStr, context);
                if( evaluationResult != null ){
                    evaluationResult = evaluationResult.toString();
                    targetCellType = CellType.FORMULA;
                }
            }else{
                evaluate(strValue, context);
            }
            if(evaluationResult == null){
                targetCellType = CellType.BLANK;
            }
        }
        return evaluationResult;
    }

    void evaluate(String strValue, Context context) {
        StringBuffer sb = new StringBuffer();
        Matcher exprMatcher = REGEX_EXPRESSION_PATTERN.matcher(strValue);
        ExpressionEvaluator evaluator = new JexlExpressionEvaluator(context.toMap());
        String matchedString;
        String expression;
        Object lastMatchEvalResult = null;
        int matchCount = 0;
        int endOffset = 0;
        while(exprMatcher.find()){
            endOffset = exprMatcher.end();
            matchCount++;
            matchedString = exprMatcher.group();
            expression = matchedString.substring(2, matchedString.length() - 1);
            lastMatchEvalResult = evaluator.evaluate(expression);
            exprMatcher.appendReplacement(sb, Matcher.quoteReplacement( lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "" ));
        }
        if( matchCount > 1 || (matchCount == 1 && endOffset < strValue.length()-1)){
            exprMatcher.appendTail(sb);
            evaluationResult = sb.toString();
        }else if(matchCount == 1){
            evaluationResult = lastMatchEvalResult;
            if(evaluationResult instanceof Number){
                targetCellType = CellType.NUMBER;
            }else if(evaluationResult instanceof Boolean){
                targetCellType = CellType.BOOLEAN;
            }
        }else if(matchCount == 0){
            evaluationResult = strValue;
        }
    }


    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    protected void updateFormulaValue() {
        if( cellType == CellType.FORMULA ){
            formula = cellValue != null ? cellValue.toString() : "";
        }else if( cellType == CellType.STRING && cellValue != null && isUserFormula(cellValue.toString())){
            formula = cellValue.toString().substring(2, cellValue.toString().length() - 1);
        }
    }

    public Pos getPos(){
        return new Pos(sheetName, row, col);
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }
    
    public Object getCellValue(){
        return cellValue;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }
    
    public boolean isFormulaCell(){
        return formula != null;
    }

    public static boolean isUserFormula(String str) {
        return str.startsWith(CellData.USER_FORMULA_PREFIX) && str.endsWith(CellData.USER_FORMULA_SUFFIX);
    }
    
    public boolean addTargetPos(Pos pos){
        return targetPos.add(pos);
    }
    
    public List<Pos> getTargetPos(){
        return targetPos;
    }

    public void resetTargetPos(){
        targetPos.clear();
    }


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof CellData)) return false;

        CellData cellData = (CellData) o;

        if (col != cellData.col) return false;
        if (row != cellData.row) return false;
        if (cellType != cellData.cellType) return false;
        if (cellValue != null ? !cellValue.equals(cellData.cellValue) : cellData.cellValue != null) return false;
        if (sheetName != null ? !sheetName.equals(cellData.sheetName) : cellData.sheetName != null) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = col;
        result = 31 * result + row;
        result = 31 * result + (sheetName != null ? sheetName.hashCode() : 0);
        result = 31 * result + (cellValue != null ? cellValue.hashCode() : 0);
        result = 31 * result + (cellType != null ? cellType.hashCode() : 0);
        return result;
    }

    @Override
    public String toString() {
        return "CellData{" +
                "sheetName=" + sheetName +
                ", row=" + row +
                ", col=" + col +
                ", cellType=" + cellType +
                ", cellValue=" + cellValue +
                '}';
    }
}
