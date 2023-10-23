import dayjs from "dayjs";
import Excel, { Cell, Style, Fill, CellRichTextValue } from "exceljs";
import {
  MergeRange,
  TemplateInfo,
  Range,
  Dimension,
  BaseAddress,
} from "./polyfill";
import { detectReplacements } from "./helper";

const getRangeDimension = (range: Range): Dimension => {
  return {
    width: range.br.col - range.tl.col + 1,
    height: range.br.row - range.tl.row + 1,
  };
};

const cloneStyleObject = ({
  numFmt,
  font,
  alignment,
  protection,
  border,
  fill,
}: Partial<Style>): Partial<Style> => {
  let fillClone: Fill | undefined;

  switch (fill?.type) {
    case "pattern":
      fillClone = {
        ...fill,
        fgColor: { ...fill.fgColor },
        bgColor: { ...fill.bgColor },
      };
      break;
    case "gradient":
      switch (fill.gradient) {
        case "angle":
          fillClone = {
            ...fill,
            stops: fill.stops?.map((stop) => ({
              position: stop.position,
              color: { ...stop.color },
            })),
          };
          break;
        case "path":
          fillClone = {
            ...fill,
            center: { ...fill.center },
            stops: fill.stops?.map((stop) => ({
              position: stop.position,
              color: { ...stop.color },
            })),
          };
          break;
      }
      break;
  }

  return {
    numFmt,
    font: {
      ...font,
      color: {
        ...font?.color,
      },
    },
    alignment: {
      ...alignment,
    },
    protection: {
      ...protection,
    },
    border: {
      top: {
        style: border?.top?.style,
        color: { ...border?.top?.color },
      },
      left: {
        style: border?.left?.style,
        color: { ...border?.left?.color },
      },
      bottom: {
        style: border?.bottom?.style,
        color: { ...border?.bottom?.color },
      },
      right: {
        style: border?.right?.style,
        color: { ...border?.right?.color },
      },
    },
    fill: fillClone,
  };
};

const applyParamsToCell = (cell: Cell, params: any) => {
  const pattern = /{{([^{}]+)}}/;
  const { value } = cell;
  if (!value) {
    return;
  }

  switch (typeof value) {
    case "string":
      {
        let cellText = new String(cell.value);
        const replaces = detectReplacements(value, pattern, params);

        replaces.forEach((replace) => {
          const [key, text] = replace;
          cellText = (text as string).replace(key, text);
        });
        cell.value = cellText.toString();
      }
      break;
    case "object":
      if (!value.hasOwnProperty("richText")) {
        return;
      }

      const cellValue = {
        richText: (value as CellRichTextValue).richText.map((t) => {
          let cellText = t.text;
          const replaces = detectReplacements(cellText, pattern, params);
          replaces.forEach((replace) => {
            const [key, text] = replace;
            cellText = (text as string).replace(key, text);
          });

          return {
            ...t,
            text: cellText,
          };
        }),
      };

      cell.value = cellValue;
      break;
  }
};

type CopyRangeInput = {
  cursor: Cell;
  templateName: string;
  onCell?: (cell: Cell) => void;
};

/**
 * @param CopyRangeInput
 * @returns range of copied cells
 */
const copyRange = ({ cursor, templateName, onCell }: CopyRangeInput): Range => {
  const wb = cursor.workbook;
  const ws = cursor.worksheet;
  const template = wb?.getWorksheet("templates");
  if ([wb, ws, template].some((l) => !l)) {
    throw new Error(
      "copyRange: workbook, worksheet or template does not existed!"
    );
  }
  ws.properties.defaultRowHeight = template.properties.defaultRowHeight;

  const matrix = wb.definedNames.getMatrix(templateName);
  const { tl, br } = matrix.getCell(
    wb.definedNames.getRanges(templateName).ranges[0]
  ) as unknown as TemplateInfo;

  /** Calculate offset */
  const offset = {
    col: cursor.fullAddress.col - tl.col,
    row: cursor.fullAddress.row - tl.row,
  };

  /** Copy template to destination */
  const styledRows = new Set();
  matrix.forEach(({ sheetName, address, row, col }) => {
    const cell = template.getCell(address);
    const targetCell = ws.getCell(row + offset.row, col + offset.col);
    /** Set row style */
    if (!styledRows.has(row)) {
      ws.getRow(Number(row + offset.row)).height = template.getRow(
        Number(row)
      ).height;
      styledRows.add(row);
    }

    /** Copy style*/
    targetCell.style = cloneStyleObject(cell.style);

    /** Copy value */
    targetCell.value = cell.value;

    /** Custom render*/
    onCell?.(targetCell);

    if (cell.model.master) {
      return;
    }

    /** Copy merged range*/
    const merge = ((template as any)._merges ?? {})[cell.address];
    if (merge) {
      const { model } = merge as MergeRange;
      ws.mergeCells(
        model.top + offset.row,
        model.left + offset.col,
        model.bottom + offset.row,
        model.right + offset.col
      );
    }
  });

  return {
    tl: {
      col: tl.col + offset.col,
      row: tl.row + offset.row,
    },
    br: {
      col: br.col + offset.col,
      row: br.row + offset.row,
    },
  };
};

export { copyRange, applyParamsToCell, getRangeDimension };
