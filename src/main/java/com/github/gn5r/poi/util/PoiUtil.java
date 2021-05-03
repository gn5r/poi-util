package com.github.gn5r.poi.util;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import com.github.gn5r.poi.util.annotation.Cell;
import com.github.gn5r.poi.util.data.CellData;
import com.github.gn5r.poi.util.data.RowData;
import com.github.gn5r.poi.util.exception.PoiNoAnnotatedException;
import com.github.gn5r.poi.util.logger.PoiUtilLogger;
import com.github.gn5r.poi.util.message.Message;
import com.github.gn5r.poi.util.message.MessageUtil;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;

/**
 *
 * @author gn5r
 */
public class PoiUtil {

	public static RowData getRowData(Workbook workbook, Sheet worksheet, Object source)
			throws IllegalArgumentException, IllegalAccessException {
		RowData rowData = new RowData();
		Set<CellData> cells = Sets.newHashSet();

		List<Field> fields = Arrays.asList(source.getClass().getDeclaredFields());
		if (isEmptyAnnotation(fields, Cell.class)) {
			throw new PoiNoAnnotatedException(source.getClass().getName(), Cell.class.getName());
		}

		for (Field f : fields) {
			f.setAccessible(true);
			if (!f.getType().equals(List.class)) {
				cells.addAll(getCellDetails(workbook, worksheet, source, f));
				rowData.setFieldName(f.getName());
			} else {
				PoiUtilLogger.warn(MessageUtil.getMessage(Message.POI0003, f.getName()));
			}
		}

		if (CollectionUtils.isNotEmpty(cells)) {
			setCells(rowData, new ArrayList<>(cells));
		}

		return rowData;
	}

	public static List<RowData> getRowDatas(Workbook workbook, Sheet worksheet, Object source)
			throws IllegalArgumentException, IllegalAccessException {
		List<RowData> rowDatas = Lists.newArrayList();

		List<Field> fields = Arrays.asList(source.getClass().getDeclaredFields());
		if (isEmptyAnnotation(fields, Cell.class)) {
			throw new PoiNoAnnotatedException(source.getClass().getName(), Cell.class.getName());
		}

		for (Field f : fields) {
			f.setAccessible(true);
			PoiUtilLogger.debug(
					MessageUtil.getMessage(Message.POI0005, f.getName(), f.getType().getSimpleName()));
			if (f.getType().equals(List.class)) {
				List<?> list = (List<?>) f.get(source);
				for (int i = 0; i < list.size(); i++) {
					Object object = list.get(i);
					RowData rowData = getRowData(workbook, worksheet, object);
					rowData.setFieldName(f.getName());
					rowData.setFirstRow(rowData.getFirstRow() + i);
					rowData.setLastRow(rowData.getLastRow() + i);
					rowDatas.add(rowData);
				}
			} else {
				Cell annotation = (Cell) f.getAnnotation(Cell.class);
				if (annotation != null) {
					Object value = f.get(source);
					RowData rowData = getRowData(workbook, worksheet, annotation, value);
					rowData.setFieldName(f.getName());
					rowDatas.add(rowData);
				} else {
					PoiUtilLogger.warn(MessageUtil.getMessage(Message.POI0001, f.getName(), Cell.class.getName()));
				}
			}
		}

		return rowDatas;
	}

	public static List<RowData> getRowDatas(Workbook workbook, Sheet worksheet, List<?> list)
			throws IllegalArgumentException, IllegalAccessException {
		List<RowData> rowDatas = Lists.newArrayList();

		PoiUtilLogger.debug(MessageUtil.getMessage(Message.POI0005, "list", list.getClass()));

		list.stream().forEach(source -> {
			List<Field> fields = Arrays.asList(source.getClass().getDeclaredFields());
			if (isEmptyAnnotation(fields, Cell.class)) {
				throw new PoiNoAnnotatedException(list.getClass().getName(), Cell.class.getName());
			}
		});

		for (int i = 0; i < list.size(); i++) {
			Object source = list.get(i);
			RowData rowData = getRowData(workbook, worksheet, source);
			rowData.setFirstRow(rowData.getFirstRow() + i);
			rowData.setLastRow(rowData.getLastRow() + i);
			rowDatas.add(rowData);
		}

		return rowDatas;
	}

	public static <T> boolean isEmptyAnnotation(List<Field> fields, T annotationClass) {
		List<Annotation[]> list = fields.stream().map(Field::getAnnotations).distinct().collect(Collectors.toList());
		Set<Annotation> annotations = Sets.newHashSet();

		for (Annotation[] arr : list) {
			List<Annotation> annotationList = Arrays.asList(arr).stream()
					.filter(a -> a.annotationType().equals(annotationClass)).distinct().collect(Collectors.toList());
			annotations.addAll(annotationList);
		}

		return CollectionUtils.isEmpty(annotations);
	}

	public static void setRowData(Workbook workbook, Sheet worksheet, RowData rowData) {
		Row row = worksheet.getRow(rowData.getFirstRow());

		if (row == null) {
			copyRow(workbook, worksheet, rowData.getFirstRow(), rowData.getFirstRow() + rowData.getHeight());
		}

		if (CollectionUtils.isNotEmpty(rowData.getCells())) {
			rowData.getCells().stream().forEach(cell -> {
				org.apache.poi.ss.usermodel.Cell newCell = row.getCell(cell.getFirstColl());
				Object value = cell.getValue();
				if (value != null) {
					if (value instanceof Number) {
						newCell.setCellValue(String.valueOf(value));
					} else {
						newCell.setCellValue(value.toString());
					}
				}
			});
		}
	}

	public static void setRowDatas(Workbook workbook, Sheet worksheet, List<RowData> rowDatas) {
		RowData source = rowDatas.get(0);

		for (int i = 0; i < rowDatas.size(); i++) {
			RowData rowData = rowDatas.get(i);
			if (rowData.getFirstRow() != null) {
				Row row = worksheet.getRow(rowData.getFirstRow());

				if (row == null) {
					row = copyRow(workbook, worksheet, source, rowData);
				}

				if (CollectionUtils.isNotEmpty(rowData.getCells())) {
					for (CellData cell : rowData.getCells()) {
						PoiUtilLogger.debug("row index:[" + i + "] " + "cell name:[" + cell.getCellName() + "]");
						org.apache.poi.ss.usermodel.Cell newCell = row.getCell(cell.getFirstColl());
						Object value = cell.getValue();
						if (value != null) {
							if (value instanceof Number) {
								newCell.setCellValue(String.valueOf(value));
							} else {
								newCell.setCellValue(value.toString());
							}
						}
					}
				}
			}
		}
	}

	public static Row copyRow(Workbook workbook, Sheet worksheet, int source, int target) {
		Row sourceRow = worksheet.getRow(source);
		Row newRow = worksheet.getRow(target);

		if (newRow != null) {
			worksheet.shiftRows(target, worksheet.getLastRowNum(), 1);
			newRow = worksheet.createRow(target);
		} else {
			newRow = worksheet.createRow(target);
		}

		// セルの型、スタイル、値などをすべてコピーする
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			org.apache.poi.ss.usermodel.Cell oldCell = sourceRow.getCell(i);
			org.apache.poi.ss.usermodel.Cell newCell = newRow.createCell(i);

			// コピー元の行が存在しない場合、処理を中断
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// スタイルのコピー
			CellStyle newCellStyle = workbook.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
			newCell.setCellStyle(newCellStyle);
		}

		// セル結合のコピー
		for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
			CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
			if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
				CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
						(newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
						cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
				worksheet.addMergedRegion(newCellRangeAddress);
			}
		}

		return newRow;
	}

	public static Row copyRow(Workbook workbook, Sheet worksheet, RowData source, RowData target) {
		Row sourceRow = worksheet.getRow(source.getFirstRow());
		Row newRow = worksheet.getRow(target.getFirstRow());

		if (newRow != null) {
			worksheet.shiftRows(target.getFirstRow(), worksheet.getLastRowNum(), target.getHeight());
			newRow = worksheet.createRow(target.getFirstRow());
		} else {
			newRow = worksheet.createRow(target.getFirstRow());
		}

		// セルの型、スタイル、値などをすべてコピーする
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			org.apache.poi.ss.usermodel.Cell oldCell = sourceRow.getCell(i);
			org.apache.poi.ss.usermodel.Cell newCell = newRow.createCell(i);

			// コピー元の行が存在しない場合、処理を中断
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// スタイルのコピー
			CellStyle newCellStyle = workbook.createCellStyle();
			newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
			newCell.setCellStyle(newCellStyle);
		}

		// セル結合のコピー
		for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
			CellRangeAddress address = worksheet.getMergedRegion(i);
			if (address.getFirstRow() == source.getFirstRow() && address.getLastRow() == source.getLastRow()) {
				CellRangeAddress newAddress = new CellRangeAddress(target.getFirstRow(), target.getLastRow(),
						address.getFirstColumn(), address.getLastColumn());
				worksheet.addMergedRegion(newAddress);
			}
		}

		return newRow;
	}

	private static List<CellData> getCellDetails(Workbook workbook, Sheet worksheet, Object source, Field f)
			throws IllegalArgumentException, IllegalAccessException {
		List<CellData> cells = Lists.newArrayList();
		Object value = f.get(source);

		Cell annotation = (Cell) f.getAnnotation(Cell.class);
		if (annotation != null) {
			Name cellName = getName(workbook, annotation.tags());
			if (cellName != null) {
				AreaReference area = new AreaReference(cellName.getRefersToFormula(),
						workbook.getSpreadsheetVersion());
				List<CellReference> cellRefs = Arrays.asList(area.getAllReferencedCells());

				if (CollectionUtils.isNotEmpty(cellRefs)) {
					cellRefs.stream().forEach(ref -> {
						CellData cell = parseToCellData(worksheet, ref, cellName, value);
						cells.add(cell);
					});
				}
			} else {
				PoiUtilLogger.warn(MessageUtil.getMessage(Message.POI0002, Arrays.toString(annotation.tags())));
			}
		} else {
			PoiUtilLogger.warn(MessageUtil.getMessage(Message.POI0001, f.getName(), Cell.class.getName()));
		}

		return cells;
	}

	private static RowData getRowData(Workbook workbook, Sheet worksheet, Cell annotation, Object value) {
		RowData rowData = new RowData();
		List<CellData> cells = getCellDetails(workbook, worksheet, annotation, value);

		if (CollectionUtils.isNotEmpty(cells)) {
			setCells(rowData, cells);
		}
		return rowData;
	}

	private static List<CellData> getCellDetails(Workbook workbook, Sheet worksheet, Cell annotation, Object value) {
		List<CellData> cells = Lists.newArrayList();
		Name cellName = getName(workbook, annotation.tags());
		if (cellName != null) {
			AreaReference area = new AreaReference(cellName.getRefersToFormula(),
					workbook.getSpreadsheetVersion());
			List<CellReference> cellRefs = Arrays.asList(area.getAllReferencedCells());

			if (CollectionUtils.isNotEmpty(cellRefs)) {
				cellRefs.stream().forEach(ref -> {
					CellData cell = parseToCellData(worksheet, ref, cellName, value);
					cells.add(cell);
				});
			}
		} else {
			PoiUtilLogger.warn(MessageUtil.getMessage(Message.POI0002, Arrays.toString(annotation.tags())));
		}
		return cells;
	}

	private static void setCells(RowData rowData, List<CellData> cells) {
		rowData.setCells(cells);
		Optional<Integer> firstRowOpt = cells.stream().filter(cell -> cell.getFirstRow() != null)
				.map(CellData::getFirstRow).distinct()
				.min(Integer::compareTo);
		Optional<Integer> lastRowOpt = cells.stream().filter(cell -> cell.getLastRow() != null)
				.map(CellData::getLastRow).distinct().max(Integer::compareTo);

		if (firstRowOpt.isPresent() && lastRowOpt.isPresent()) {
			int height = (lastRowOpt.get() - firstRowOpt.get()) + 1;
			rowData.setFirstRow(firstRowOpt.get());
			rowData.setLastRow(lastRowOpt.get());
			rowData.setHeight(height);
		}
	}

	private static CellData parseToCellData(Sheet worksheet, CellReference ref, Name name, Object cellValue) {
		CellData cell = new CellData();
		cell.setFirstRow(ref.getRow());
		cell.setLastRow(ref.getRow());
		cell.setFirstColl(ref.getCol());
		cell.setLastColl(ref.getCol());
		cell.setCellName(name.getNameName());
		cell.setValue(cellValue);

		for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
			CellRangeAddress rangeAddress = worksheet.getMergedRegion(i);
			if (rangeAddress.getFirstRow() == ref.getRow() && rangeAddress.getFirstColumn() == ref.getCol()) {
				cell.setFirstRow(rangeAddress.getFirstRow());
				cell.setLastRow(rangeAddress.getLastRow());
				cell.setFirstColl(rangeAddress.getFirstColumn());
				cell.setLastColl(rangeAddress.getLastColumn());
			}
		}

		return cell;
	}

	private static Name getName(Workbook workbook, String[] tags) {
		for (String tag : tags) {
			Name name = workbook.getName(tag);
			if (name != null) {
				return name;
			}
		}
		return null;
	}
}
