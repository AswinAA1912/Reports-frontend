import dayjs from "dayjs";

export const mapForExport = (
  columns: any[],
  rows: any[]
) => {
  const headers = columns.map(c => c.header || c.label);
  const data = rows.map((row, index) =>
    columns.map(col => {
      if (col.key === "sno") return index + 1;

      let value = row[col.key];
      if ((value === undefined || value === null) && col.altKey) {
        value = row[col.altKey];
      }

      if (col.type === "date" && value) {
        return dayjs(value).format("DD/MM/YYYY");
      }

      return value ?? "";
    })
  );

  return { headers, data };
};
