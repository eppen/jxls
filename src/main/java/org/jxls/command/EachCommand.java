package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.GroupData;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Collection;
import java.util.List;


/**
 * Implements iteration over collection of items
 * 'items' is a bean name of the collection in context
 * 'var' is a name of a collection item to put into the context during the iteration
 * 'direction' defines expansion by rows (DOWN) or by columns (RIGHT). Default is DOWN.
 * 'cellRefGenerator' defines custom strategy for target cell references.
 * Date: Nov 10, 2009
 *
 * @author Leonid Vysochyn
 */
public class EachCommand extends AbstractCommand {

    public static final String COMMAND_NAME = "each";

    public enum Direction {RIGHT, DOWN}

    static final String GROUP_DATA_KEY = "_group";

    private String var;
    private String items;
    private String select;
    private Area area;
    private Direction direction = Direction.DOWN;
    private CellRefGenerator cellRefGenerator;
    private String multisheet;
    private String groupBy;
    private String groupOrder;

    private static Logger logger = LoggerFactory.getLogger(EachCommand.class);

    public EachCommand() {
    }

    /**
     * @param var       name of the key in the context to contain each collection items during iteration
     * @param items     name of the collection bean in the context
     * @param direction defines processing by rows (DOWN - default) or columns (RIGHT)
     */
    public EachCommand(String var, String items, Direction direction) {
        this.var = var;
        this.items = items;
        this.direction = direction == null ? Direction.DOWN : direction;
    }

    public EachCommand(String items, Area area) {
        this(null, items, area);
    }

    public EachCommand(String var, String items, Area area) {
        this(var, items, area, Direction.DOWN);
    }

    public EachCommand(String var, String items, Area area, Direction direction) {
        this(var, items, direction);
        if (area != null) {
            this.area = area;
            addArea(this.area);
        }
    }

    /**
     * @param var              name of the key in the context to contain each collection items during iteration
     * @param items            name of the collection bean in the context
     * @param area             body area for this command
     * @param cellRefGenerator generates target cell ref for each collection item during iteration
     */
    public EachCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
        this(var, items, area, (Direction) null);
        this.cellRefGenerator = cellRefGenerator;
    }

    /**
     * Gets iteration directino
     *
     * @return current direction for iteration
     */
    public Direction getDirection() {
        return direction;
    }

    /**
     * Sets iteration direction
     *
     * @param direction
     */
    public void setDirection(Direction direction) {
        this.direction = direction;
    }

    public void setDirection(String direction) {
        this.direction = Direction.valueOf(direction);
    }

    /**
     * Gets defined cell ref generator
     *
     * @return current {@link CellRefGenerator} instance or null
     */
    public CellRefGenerator getCellRefGenerator() {
        return cellRefGenerator;
    }

    public void setCellRefGenerator(CellRefGenerator cellRefGenerator) {
        this.cellRefGenerator = cellRefGenerator;
    }

    public String getName() {
        return COMMAND_NAME;
    }

    /**
     * Gets current variable name for collection item in the context during iteration
     *
     * @return collection item key name in the context
     */
    public String getVar() {
        return var;
    }

    /**
     * Sets current variable name for collection item in the context during iteration
     *
     * @param var
     */
    public void setVar(String var) {
        this.var = var;
    }

    /**
     * Gets collection bean name
     *
     * @return collection bean name in the context
     */
    public String getItems() {
        return items;
    }

    /**
     * Sets collection bean name
     *
     * @param items collection bean name in the context
     */
    public void setItems(String items) {
        this.items = items;
    }

    /**
     * Gets current 'select' expression for filtering out collection items
     *
     * @return current 'select' expression or null if undefined
     */
    public String getSelect() {
        return select;
    }

    /**
     * Sets current 'select' expression for filtering collection
     *
     * @param select filtering expression
     */
    public void setSelect(String select) {
        this.select = select;
    }

    /**
     * @return Context variable name holding a list of Excel sheet names to output the collection to
     */
    public String getMultisheet() {
        return multisheet;
    }

    /**
     * Sets name of context variable holding a list of Excel sheet names to output the collection to
     *
     * @param multisheet
     */
    public void setMultisheet(String multisheet) {
        this.multisheet = multisheet;
    }

    /**
     * @return property name for grouping the collection
     */
    public String getGroupBy() {
        return groupBy;
    }

    /**
     * @param groupBy property name for grouping the collection
     */
    public void setGroupBy(String groupBy) {
        this.groupBy = groupBy;
    }

    /**
     * @return group order
     */
    public String getGroupOrder() {
        return groupOrder;
    }

    /**
     * @param groupOrder group ordering
     */
    public void setGroupOrder(String groupOrder) {
        this.groupOrder = groupOrder;
    }

    @Override
    public Command addArea(Area area) {
        if (area == null) {
            return this;
        }
        if (areaList.size() >= 1) {
            throw new IllegalArgumentException("You can add only a single area to 'each' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Collection itemsCollection = Util.transformToCollectionObject(getTransformationConfig().getExpressionEvaluator(), items, context);
        if (groupBy == null || groupBy.length() == 0) {
            return processCollection(context, itemsCollection, cellRef, var);
        } else {
            Collection<GroupData> groupedData = Util.groupCollection(itemsCollection, groupBy, groupOrder);
            String groupVar = var != null ? var : GROUP_DATA_KEY;
            return processCollection(context, groupedData, cellRef, groupVar);
        }
    }

    private Size processCollection(Context context, Collection itemsCollection, CellRef cellRef, String varName) {
        int index = 0;
        int newWidth = 0;
        int newHeight = 0;

        CellRefGenerator cellRefGenerator = this.cellRefGenerator;
        if (cellRefGenerator == null && multisheet != null) {
            List<String> sheetNameList = extractSheetNameList(context);
            cellRefGenerator = new SheetNameGenerator(sheetNameList, cellRef);
        }
        CellRef currentCell = cellRefGenerator != null ? cellRefGenerator.generateCellRef(index, context) : cellRef;
        JexlExpressionEvaluator selectEvaluator = null;
        if (select != null) {
            selectEvaluator = new JexlExpressionEvaluator(select);
        }

        for (Object obj : itemsCollection) {
            context.putVar(varName, obj);
            if (selectEvaluator != null && !Util.isConditionTrue(selectEvaluator, context)) {
                context.removeVar(varName);
                continue;
            }
            Size size = area.applyAt(currentCell, context);
            index++;
            if (cellRefGenerator != null) {
                newWidth = Math.max(newWidth, size.getWidth());
                newHeight = Math.max(newHeight, size.getHeight());
                if (index < itemsCollection.size()) {
                    currentCell = cellRefGenerator.generateCellRef(index, context);
                }
            } else if (direction == Direction.DOWN) {
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow() + size.getHeight(), currentCell.getCol());
                newWidth = Math.max(newWidth, size.getWidth());
                newHeight += size.getHeight();
            } else {
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                newWidth += size.getWidth();
                newHeight = Math.max(newHeight, size.getHeight());
            }
            context.removeVar(varName);
        }
        return new Size(newWidth, newHeight);
    }

    private List<String> extractSheetNameList(Context context) {
        try {
            return (List<String>) context.getVar(multisheet);
        } catch (Exception e) {
            throw new JxlsException("Failed to get sheet names from " + multisheet, e);
        }
    }

}
