import dayjs from "dayjs";
import Excel from "exceljs";
import {
  applyParamsToCell,
  copyRange,
  getRangeDimension,
} from "./excel-lobster";
import { BaseAddress } from "./polyfill";
import fs from "fs";

type TemplateParams = {
  docSequence: string;
  requestedAt: string;
  requester: string;
  teamCode: string;
  teamName: string;
  companyName: string;
  accountNumber: string;
  beneficiaryName: string;
  bankName: string;
  branchName: string;
  bankCode: string;
  totalAmountBeforeTax: string;
  totalAmount: string;
  costCenter: string;
  expenses: {
    order: number;
    note: string;
    totalAmountWithoutVat: string;
    totalVatAmount: string;
    totalAmountWithVat: string;
  }[];
};

const exec = async () => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("./resources/ReportTemplate.xlsx");

  const ws = workbook.getWorksheet("Sheet");
  const pageX = {
    startCol: 2,
    endCol: 18,
  };
  let anchor: BaseAddress;
  /** e.g: render block 1 */
  {
    const range = copyRange({
      cursor: ws.getCell("B2"),
      templateName: "section1",
      onCell: (cell) => {
        applyParamsToCell(cell, {
          docSequence: "RP001",
          requestedAt: dayjs().format("DD/MM/YYYY") ?? "",
          requester: "Ngô Đăng Khôi",
          teamCode: "team-001",
          teamName: "the coding gangz",
          costCenter: "CT-001",
          companyName: "Công",
          accountNumber: "04001010192938",
          beneficiaryName: "Lê Quang Móm",
          bankName: "Maritime bank",
          branchName: "Thủ Đức",
          bankCode: "MT-001",
        });
      },
    });

    anchor = { col: range.tl.col, row: range.br.row + 1 };
  }

  /** e.g: render block 2 */
  {
    const expenses = [
      {
        order: "1",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "2",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "3",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "4",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "5",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "6",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "7",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "8",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
      {
        order: "9",
        note: "Lorem Ipsum is simply dummy text of the printing and typesetting",
        totalAmountWithoutVat: "10,000,000",
        totalVatAmount: "10,000,000",
        totalAmountWithVat: "10,000,000",
      },
    ];

    expenses.forEach((expense) => {
      const range = copyRange({
        cursor: ws.getCell(anchor.row, anchor.col),
        templateName: "section2",
        onCell: (cell) => {
          applyParamsToCell(cell, expense);
        },
      });

      anchor = { col: range.tl.col, row: range.br.row + 1 };
    });
  }

  /** e.g: render block 3 */
  {
    const range = copyRange({
      cursor: ws.getCell(anchor.row, anchor.col),
      templateName: "section3",
      onCell: (cell) => {
        applyParamsToCell(cell, {
          totalAmountBeforeTax: "1,000,000",
          totalAmount: "1,000,000",
        });
      },
    });

    anchor = { col: range.tl.col, row: range.br.row + 1 };
  }

  /** e.g: render block 4 */
  {
    const range = copyRange({
      cursor: ws.getCell(anchor.row, anchor.col),
      templateName: "section4",
      onCell: (cell) => {
        if (cell.value === "{{isApproved}}") {
          /** you have to handle individual cell in case of special logic e.g: add image to cell, conditional fill pattern, etc.*/
          cell.value = "";
          return;
        }
        applyParamsToCell(cell, {
          approver: "Trần Việt Khoa",
          approveDate: "12/8/2023",
        });
      },
    });
    anchor = { col: range.br.col + 1, row: range.tl.row };
  }

  /** e.g: render block 5 */
  {
    const listOfApprover = [
      {
        approver: "Vũ Đức Nhân",
        approveDate: "13/8/2023",
      },
      {
        approver: "Trần Thanh Huy",
        approveDate: "14/8/2023",
      },
      {
        approver: "Nguyễn Vũ Thuần",
        approveDate: "20/8/2023",
      },
      {
        approver: "Ngô Đức Kế",
        approveDate: "01/9/2023",
      },
      {
        approver: "Nguyễn Ngọc Ngân",
        approveDate: "01/9/2023",
      },
      {
        approver: "Trần Thanh Ngân",
        approveDate: "02/9/2023",
      },
      {
        approver: "Trần Ngọc Thư",
        approveDate: "02/9/2023",
      },
      {
        approver: "Lê Quang Hiếu",
        approveDate: "02/9/2023",
      },
      {
        approver: "Trần Thanh Ngân",
        approveDate: "02/9/2023",
      },
      {
        approver: "Trần Ngọc Thư",
        approveDate: "02/9/2023",
      },
      {
        approver: "Lê Quang Hiếu",
        approveDate: "02/9/2023",
      },
    ];

    listOfApprover.forEach((approver, index) => {
      const range = copyRange({
        cursor: ws.getCell(anchor.row, anchor.col),
        templateName:
          index === 6 || (index > 6 && index / 7 === 0)
            ? "section6"
            : "section5",
        onCell: (cell) => {
          if (cell.value === "{{isApproved}}") {
            /** you have to handle individual cell in case of special logic e.g: add image to cell, conditional fill pattern, etc.*/
            return;
          }
          applyParamsToCell(cell, approver);
        },
      });

      const { width } = getRangeDimension(range);

      /** auto wrap*/
      if (range.br.col + width > pageX.endCol) {
        anchor = {
          row: range.br.row + 1,
          col: pageX.startCol,
        };
      } else {
        anchor = { col: range.br.col + 1, row: range.tl.row };
      }
    });
  }

  const path = "./output";
  if (!fs.existsSync(path)) {
    fs.mkdirSync(path);
  }

  await workbook.xlsx.writeFile(
    `${path}/rp_${dayjs().format("DD_MM_HHmmss")}.xlsx`
  );
  console.log("done...............");
};

exec();
