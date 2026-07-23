import * as React from "react";
import { sp } from "../DataUploadWebPart";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import {
  TextField,
  DefaultButton,
  Dropdown,
  IDropdownOption,
} from "@fluentui/react";

import { useState, useEffect } from "react";
import Left from "../assets/LeftArrow.png";
import Right from "../assets/RightArrow.png";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface IData {
  Id: number;
  Username: string;
  Created: string;
  Department: string;
  VendorName: string;
  VendorCode: string;
  PONumber: string;
  GLCode: string;
  GLDescription: string;
  EmployeeCostCenter: string;
  EmployeeCostCenterName: string;
  Amount: number;
  ExpenseMonth: string;
  Remarks: string;
  Status?: string;
}

export default function AdjustmentReport() {
  const [isPerformer, setIsPerformer] = React.useState(false);
  const [file, setFile] = React.useState<File | null>(null);
  const [data, setData] = React.useState<IData[]>([]);
  const [isSearched, setIsSearched] = React.useState(false);
  const [closedMonths, setClosedMonths] = React.useState<string[]>([]);
  const [editingIds, setEditingIds] = useState<number[]>([]);
  const [userName, setUserName] = React.useState("");
  const [department, setDepartment] = React.useState("");
  const [vendorName, setVendorName] = React.useState("");
  const [poNumber, setPoNumber] = React.useState("");
  const [fromDate, setFromDate] = React.useState("");
  const [toDate, setToDate] = React.useState("");
  const [userOptions, setUserOptions] = React.useState<string[]>([]);
  const [deptOptions, setDeptOptions] = React.useState<string[]>([]);
  const [vendorOptions, setVendorOptions] = React.useState<string[]>([]);
  const [poOptions, setPoOptions] = React.useState<string[]>([]);

  const [filteredData, setFilteredData] = useState<any[]>([]);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);
  const totalPages = Math.ceil(data.length / itemsPerPage);
  const [selectedIds, setSelectedIds] = useState<number[]>([]);
  //const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const handleExit = () => {
    //window.location.href = `${window.location.origin}/sites/SonaFinance/SitePages/Accuralsheet.aspx`;
    window.location.href = `https://sonacomstargroup.sharepoint.com/sites/RLY_Finance_UAT/SitePages/Accuralsheet.aspx`;
  };

  const searchData = async () => {
    let filter: string[] = [];

    // ✅ Hide deleted
    filter.push("DeleteFlag ne 1");

    // ✅ Only Pending
    filter.push("(Status eq 'Pending' or Status eq 'pending')");

    // ✅ Text filters
    if (userName) filter.push(`substringof('${userName}', Username)`);
    if (department) filter.push(`substringof('${department}', Department)`);
    if (vendorName) filter.push(`substringof('${vendorName}', VendorName)`);
    if (poNumber) filter.push(`substringof('${poNumber}', PONumber)`);

    // ✅ Date filters
    if (fromDate) {
      filter.push(`Created ge datetime'${new Date(fromDate).toISOString()}'`);
    }

    if (toDate) {
      const to = new Date(toDate);
      to.setDate(to.getDate() + 1);
      filter.push(`Created lt datetime'${to.toISOString()}'`);
    }

    const query = filter.join(" and ");

    try {
      // ✅ STEP 1: GET MAIN DATA
      const items = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.filter(query)
        .top(5000)();

      console.log("All Items:", items);

      // ✅ STEP 2: GET CLOSED MONTHS (Month + Year)
      const closedItems = await sp.web.lists
        .getByTitle("PerformerClosureAccess")
        .items.select("Month", "DateofClosure")
        .top(5000)();

      const closedMonths = closedItems.map((i: any) => {
        const month = String(i.Month || "")
          .trim()
          .toLowerCase();
        const year = i.DateofClosure
          ? new Date(i.DateofClosure).getFullYear()
          : "";

        return `${month}-${year}`;
      });

      console.log("Closed Months:", closedMonths);

      // ✅ STEP 3: FORMAT MONTH
      const formatMonth = (value: any) => {
        if (!value) return "";
        const str = String(value).trim().toLowerCase();
        return str.charAt(0).toUpperCase() + str.slice(1);
      };

      // ✅ STEP 4: FINAL FILTER (FIXED)
      const filteredData = items
        .filter((item: any) => {
          const month = String(item.ExpenseMonth || "")
            .trim()
            .toLowerCase();

          // ✅ Use Created Year (or your custom Year field if exists)
          const year = item.Created ? new Date(item.Created).getFullYear() : "";

          const monthYear = `${month}-${year}`;

          const status = String(item.Status || "")
            .trim()
            .toLowerCase();

          // ✅ Only Pending
          if (status !== "pending") return false;

          // ❌ Hide closed month-year only
          if (closedMonths.includes(monthYear)) return false;

          return true;
        })
        .map((item: any) => ({
          ...item,
          ExpenseMonth: formatMonth(item.ExpenseMonth),
        }));

      console.log("Final Data:", filteredData);

      setIsSearched(true);
      setData(filteredData);
      setFilteredData(filteredData); // ✅ for pagination
    } catch (error) {
      console.error("Search Error:", error);
      alert("Error fetching data");
    }
  };

  const removeNewRow = (index: number) => {
    const updated = [...data];
    updated.splice(index, 1); // remove row by index
    setData(updated);
  };
  const getClosureData = async () => {
    const items = await sp.web.lists
      .getByTitle("PerformerClosureAccess")
      .items.select("Month", "DateofClosure")
      .top(5000)();

    return items.map((item: any) => ({
      month: item.Month?.toLowerCase(),
      closingDate: new Date(item.DateofClosure),
    }));
  };

  const deleteSelectedItems = async () => {
    if (selectedIds.length === 0) {
      alert("Please select records");
      return;
    }

    const confirmDelete = window.confirm(
      `Are you sure you want to delete ${selectedIds.length} record(s)?`,
    );

    if (!confirmDelete) return;

    try {
      for (const id of selectedIds) {
        await sp.web.lists
          .getByTitle("AccrualSheetList")
          .items.getById(id)
          .update({
            DeleteFlag: true,
          });
      }

      alert("Selected records deleted successfully");

      setSelectedIds([]);

      await searchData();
    } catch (error) {
      console.log(error);
      alert("Error deleting records");
    }
  };

  const updateGLCode = async (item: IData) => {
    try {
      await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.getById(item.Id)
        .update({
          GLCode: item.GLCode,
        });

      alert("GL Code updated successfully");

      setEditingIds(editingIds.filter((id) => id !== item.Id));

      await searchData();
    } catch (error) {
      console.log(error);
      alert("Error updating GL Code");
    }
  };
  const getClosedMonths = async () => {
    const items = await sp.web.lists
      .getByTitle("PerformerClosureAccess")
      .items.select("Month", "DateofClosure")
      .top(5000)();

    // ✅ Store Month + Year
    return items.map((i: any) => {
      const month = String(i.Month || "")
        .trim()
        .toLowerCase();
      const year = i.DateofClosure
        ? new Date(i.DateofClosure).getFullYear()
        : "";

      return `${month}-${year}`;
    });
  };

  
  const searchData1 = async () => {
    debugger;

    let filter: string[] = [];

    // ✅ Default filters
    filter.push("DeleteFlag ne 1");

    // ✅ SAFE STATUS FILTER (important fix)
    filter.push("(Status eq 'Pending' or Status eq 'pending')");

    // ✅ Text filters
    if (userName && userName.trim() !== "") {
      filter.push(`substringof('${userName.trim()}', Username)`);
    }

    if (department && department.trim() !== "") {
      filter.push(`substringof('${department.trim()}', Department)`);
    }

    if (vendorName && vendorName.trim() !== "") {
      filter.push(`substringof('${vendorName.trim()}', VendorName)`);
    }

    if (poNumber && poNumber.trim() !== "") {
      filter.push(`substringof('${poNumber.trim()}', PONumber)`);
    }

    // ✅ Date filters (FIXED)
    if (fromDate && !isNaN(new Date(fromDate).getTime())) {
      filter.push(`Created ge datetime'${new Date(fromDate).toISOString()}'`);
    }

    if (toDate && !isNaN(new Date(toDate).getTime())) {
      const to = new Date(toDate);
      to.setDate(to.getDate() + 1); // ✅ IMPORTANT FIX
      filter.push(`Created lt datetime'${to.toISOString()}'`);
    }

    const query = filter.join(" and ");

    try {
      // ✅ STEP 1: FETCH MAIN DATA
      const items = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.filter(query)
        .top(5000)();

      console.log("All Items:", items);

      // ✅ STEP 2: FETCH CLOSED MONTHS
      const closedItems = await sp.web.lists
        .getByTitle("PerformerClosureAccess")
        .items.select("Month")
        .top(5000)();

      const closedMonths = closedItems.map((i: any) =>
        String(i.Month || "")
          .trim()
          .toLowerCase(),
      );

      console.log("Closed Months:", closedMonths);

      // ✅ STEP 3: FORMAT MONTH (April)
      const formatMonth = (value: any) => {
        if (!value) return "";

        const str = String(value).trim().toLowerCase();
        return str.charAt(0).toUpperCase() + str.slice(1);
      };

      // ✅ STEP 4: FINAL FILTER LOGIC 🔥
      const filteredData = items
        .filter((item: any) => {
          const month = String(item.ExpenseMonth || "")
            .trim()
            .toLowerCase();

          const status = String(item.Status || "")
            .trim()
            .toLowerCase();

          console.log("Checking → Month:", month, "Status:", status);

          // ✅ Ensure pending
          if (status !== "pending") return false;

          // ❌ Hide closed months ONLY
          return !closedMonths.includes(month);
        })
        .map((item: any) => ({
          ...item,
          ExpenseMonth: formatMonth(item.ExpenseMonth), // ✅ Show "April"
        }));

      console.log("Final Data:", filteredData);

      // ✅ SET DATA
      setIsSearched(true);
      setData(filteredData);
    } catch (error) {
      console.error("Search Error:", error);
      alert("Error fetching data");
    }
  };

  const formatMonth = (value: any) => {
    if (!value) return "";

    // If it's a valid date → convert to month name
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleString("en-US", { month: "long" });
    }

    // If already string → capitalize properly
    const str = String(value).trim().toLowerCase();
    return str.charAt(0).toUpperCase() + str.slice(1);
  };

  // const saveAllRows = async () => {
  //   const newRows = data.filter((item) => item.Id === 0);

  //   if (newRows.length === 0) {
  //     alert("No new rows to save");
  //     return;
  //   }

  //   // ✅ Helper: month name → number
  //   const getMonthNumber = (monthName: string) => {
  //     const months: any = {
  //       january: 0,
  //       february: 1,
  //       march: 2,
  //       april: 3,
  //       may: 4,
  //       june: 5,
  //       july: 6,
  //       august: 7,
  //       september: 8,
  //       october: 9,
  //       november: 10,
  //       december: 11,
  //     };
  //     return months[monthName.toLowerCase()];
  //   };

  //   const today = new Date();
  //   const currentMonth = today.getMonth();
  //   const currentYear = today.getFullYear();
  //   const currentDate = today.getDate();

  //   let skippedUsers: string[] = [];

  //   try {
  //     for (const item of newRows) {
  //       const username = item.Username || "Unknown";

  //       // ✅ Required validation
  //       if (
  //         !item.Username ||
  //         !item.Department ||
  //         !item.VendorName ||
  //         !item.VendorCode ||
  //         !item.PONumber ||
  //         !item.GLCode ||
  //         !item.GLDescription ||
  //         !item.EmployeeCostCenter ||
  //         !item.EmployeeCostCenterName ||
  //         !item.Amount ||
  //         !item.ExpenseMonth ||
  //         !item.Remarks
  //       ) {
  //         alert("All fields are mandatory");
  //         return;
  //       }

  //       // ✅ Amount validation
  //       if (isNaN(Number(item.Amount))) {
  //         alert("Amount must be numeric");
  //         return;
  //       }

  //       // ✅ Month validation
  //       const expenseMonthStr = String(item.ExpenseMonth).trim().toLowerCase();

  //       const expMonth = getMonthNumber(expenseMonthStr);

  //       if (expMonth === undefined) {
  //         skippedUsers.push(username);
  //         continue;
  //       }

  //       const expYear = currentYear;
  //       let isValid = false;

  //       // ✅ Current month
  //       if (expMonth === currentMonth) {
  //         isValid = true;
  //       }

  //       // ✅ Previous month till 5th
  //       const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  //       const prevYear = currentMonth === 0 ? currentYear - 1 : currentYear;

  //       if (
  //         expMonth === prevMonth &&
  //         expYear === prevYear &&
  //         currentDate <= 5
  //       ) {
  //         isValid = true;
  //       }

  //       // ✅ NEW: Allow future months
  //       if (expMonth > currentMonth) {
  //         isValid = true;
  //       }

  //       // ❌ Skip invalid
  //       if (!isValid) {
  //         skippedUsers.push(username);
  //         continue;
  //       }

  //       // ✅ Insert
  //       await sp.web.lists.getByTitle("AccrualSheetList").items.add({
  //         Username: item.Username,
  //         Department: item.Department,
  //         VendorName: item.VendorName,
  //         VendorCode: item.VendorCode,
  //         PONumber: item.PONumber,
  //         GLCode: item.GLCode,
  //         GLDescription: item.GLDescription,
  //         EmployeeCostCenter: item.EmployeeCostCenter,
  //         EmployeeCostCenterName: item.EmployeeCostCenterName,
  //         Amount: Number(item.Amount),
  //         ExpenseMonth: expenseMonthStr,
  //         Remarks: item.Remarks,
  //         DeleteFlag: false,
  //         Status: "Pending",
  //       });
  //     }

  //     // ✅ Remove duplicate usernames (ES5 safe)
  //     const uniqueSkippedUsers = skippedUsers.filter((item, index) => {
  //       return skippedUsers.indexOf(item) === index;
  //     });

  //     // ✅ Final message
  //     if (uniqueSkippedUsers.length > 0) {
  //       alert(
  //         `Upload complete ✅\nSkipped users: ${uniqueSkippedUsers.join(", ")}`,
  //       );
  //     } else {
  //       alert("All records saved successfully ✅");
  //     }

  //     void searchData();
  //   } catch (error) {
  //     console.log("Bulk save error:", error);
  //     alert("Error saving records");
  //   }
  // };
  const saveAllRows = async () => {
    const newRows = data.filter((item) => item.Id === 0);

    if (newRows.length === 0) {
      alert("No new rows to save");
      return;
    }

    const getMonthNumber = (monthName: string) => {
      const months: any = {
        january: 0,
        february: 1,
        march: 2,
        april: 3,
        may: 4,
        june: 5,
        july: 6,
        august: 7,
        september: 8,
        october: 9,
        november: 10,
        december: 11,
      };
      return months[monthName.toLowerCase()];
    };

    const capitalizeMonth = (month: string) => {
      const m = month.trim().toLowerCase();
      return m.charAt(0).toUpperCase() + m.slice(1);
    };

    const today = new Date();
    const currentMonth = today.getMonth();
    const currentDate = today.getDate();

    let errorList: string[] = [];
    let errorRows: any[] = []; // ✅ store full row

    // ===============================
    // ✅ STEP 1: VALIDATE ALL ROWS
    // ===============================
    for (const item of newRows) {
      const username = item.Username || "Unknown";

      // Required validation
      if (
        !item.Username ||
        !item.Department ||
        !item.VendorName ||
        !item.VendorCode ||
        !item.PONumber ||
        !item.GLCode ||
        !item.GLDescription ||
        !item.EmployeeCostCenter ||
        !item.EmployeeCostCenterName ||
        !item.Amount ||
        !item.ExpenseMonth ||
        !item.Remarks
      ) {
        errorList.push(`${username} → Missing required fields`);
        errorRows.push(item);
        continue;
      }

      // Amount validation
      if (isNaN(Number(item.Amount))) {
        errorList.push(`${username} → Amount must be numeric`);
        errorRows.push(item);
        continue;
      }

      if (Number(item.Amount) < 0) {
        errorList.push(`${username} → Negative amount not allowed`);
        errorRows.push(item);
        continue;
      }

      // Month validation
      const expenseMonthStr = String(item.ExpenseMonth || "")
        .trim()
        .toLowerCase();

      const expMonth = getMonthNumber(expenseMonthStr);

      if (expMonth === undefined) {
        errorList.push(`${username} → Invalid month`);
        errorRows.push(item);
        continue;
      }

      let isValid = false;

      // ✅ Current month
      if (expMonth === currentMonth) {
        isValid = true;
      }

      // ✅ Previous month till 5th
      const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;

      if (expMonth === prevMonth && currentDate <= 5) {
        isValid = true;
      }

      // ✅ Future month
      if (expMonth > currentMonth) {
        isValid = true;
      }

      // ❌ Past month
      if (!isValid) {
        errorList.push(
          `${username}-${expenseMonthStr} → Past month not allowed`,
        );
        errorRows.push(item);
      }
    }

    // ===============================
    // ❌ STEP 2: STOP IF ERROR
    // ===============================
    if (errorList.length > 0) {
      alert(
        "❌ Some records are invalid. Nothing saved.\n\n" +
          errorList.join("\n"),
      );

      // ✅ SHOW ERROR ROWS IN GRID
      setData(errorRows);
      setFilteredData(errorRows);
      setIsSearched(true);

      return;
    }

    // ===============================
    // ✅ STEP 3: SAVE ALL ROWS
    // ===============================
    try {
      for (const item of newRows) {
        const expenseMonthStr = capitalizeMonth(item.ExpenseMonth);

        await sp.web.lists.getByTitle("AccrualSheetList").items.add({
          Title: String(item.Username || ""),

          Username: String(item.Username || ""),
          Department: String(item.Department || ""),
          VendorName: String(item.VendorName || ""),
          VendorCode: String(item.VendorCode || ""),
          PONumber: String(item.PONumber || ""),
          GLCode: String(item.GLCode || ""),
          GLDescription: String(item.GLDescription || ""),
          EmployeeCostCenter: String(item.EmployeeCostCenter || ""),
          EmployeeCostCenterName: String(item.EmployeeCostCenterName || ""),
          Amount: Number(item.Amount || 0),

          // ✅ IMPORTANT (fix SharePoint error)
          ExpenseMonth: expenseMonthStr,

          Remarks: String(item.Remarks || ""),
          DeleteFlag: false,
          Status: "Pending",
        });
      }

      alert("All records saved successfully ✅");

      // ✅ Clear error grid
      setData([]);
      setFilteredData([]);
      setIsSearched(false);

      // Reload
      void searchData();
    } catch (error) {
      console.log("Bulk save error:", error);
      alert("Error saving records");
    }
  };

  // SEARCH DATA
  const loadDropdownData = async () => {
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.select("Username", "Department", "VendorName", "PONumber")
        .top(5000)();

      // ✅ FIXED unique function (NO Set, NO spread)
      const unique = (arr: any[], key: string) => {
        const result: string[] = [];

        arr.forEach((item) => {
          const value = item[key];

          if (value && result.indexOf(value) === -1) {
            result.push(value);
          }
        });

        return result;
      };

      setUserOptions(unique(items, "Username"));
      setDeptOptions(unique(items, "Department"));
      setVendorOptions(unique(items, "VendorName"));
      setPoOptions(unique(items, "PONumber"));
    } catch (error) {
      console.log("Dropdown load error:", error);
    }
  };

  const freezeData = async () => {
    if (data.length === 0) {
      alert("No records to freeze");
      return;
    }

    try {
      // Filter only Pending records
      const pendingItems = data.filter((item) => item.Status === "Pending");

      if (pendingItems.length === 0) {
        alert("No Pending records found");
        return;
      }

      // 🔁 Update all pending items
      for (const item of pendingItems) {
        await sp.web.lists
          .getByTitle("AccrualSheetList")
          .items.getById(item.Id)
          .update({
            Status: "Freez",
          });
      }

      alert("Records freeze successfully ✅");

      // 🔄 Refresh data
      void searchData(); // or your fetch function
    } catch (error) {
      console.log("Freeze error:", error);
      alert("Error freezing records");
    }
  };

  const updateRow = (index: number, field: string, value: any) => {
    const updated = [...data];

    updated[index] = {
      ...updated[index],
      [field]: value,
    };

    setData(updated);
  };
  const checkUserAccess = async () => {
    try {
      const user = await sp.web.currentUser();

      const groups = await sp.web.siteUsers.getById(user.Id).groups();

      const performer = groups.some((g: any) => g.Title === "AccrualPerformer");

      setIsPerformer(performer);
    } catch (error) {
      console.log("Access check error", error);
    }
  };

  const exportExcel = () => {
    const exportData = data.map((item: any) => ({
      UserName: item.Username,
      Department: item.Department,
      VendorName: item.VendorName,
      VendorCode: item.VendorCode,
      PONumber: item.PONumber,
      GLCode: item.GLCode,
      GLDescription: item.GLDescription,
      EmployeeCostCenter: item.EmployeeCostCenter,
      EmployeeCostCenterName: item.EmployeeCostCenterName,
      Amount: item.Amount,
      ExpenseMonth: item.ExpenseMonth,
      Remarks: item.Remarks,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    const workbook = {
      Sheets: { "Accrual Report": worksheet },
      SheetNames: ["Accrual Report"],
    };

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "Accrual_Report.xlsx");
  };

  // DELETE ROW
  const deleteItem = async (id: number) => {
    // ✅ Confirmation popup
    const confirmDelete = window.confirm(
      "Are you sure you want to delete the data?",
    );

    // ❌ If user clicks Cancel → stop
    if (!confirmDelete) {
      return;
    }

    try {
      await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.getById(id)
        .update({
          DeleteFlag: true,
        });

      // ✅ Success message
      alert("Data deleted successfully");

      void searchData();
    } catch (error) {
      console.log("Delete error:", error);
      alert("Error deleting data");
    }
  };

  const addNewRow = () => {
    const newRow: IData = {
      Id: 0,
      Username: "",
      Created: "",
      Department: "",
      VendorName: "",
      VendorCode: "",
      PONumber: "",
      GLCode: "",
      GLDescription: "",
      EmployeeCostCenter: "",
      EmployeeCostCenterName: "",
      Amount: 0,
      ExpenseMonth: "",
      Remarks: "",
    };

    const updated = [...data, newRow];

    setData(updated);

    // Move to last page
    setCurrentPage(Math.ceil(updated.length / itemsPerPage));
  };

  const saveRow = async (item: IData) => {
    // Validation
    if (
      !item.Username ||
      !item.Department ||
      !item.VendorName ||
      !item.VendorCode ||
      !item.PONumber ||
      !item.GLCode ||
      !item.GLDescription ||
      !item.EmployeeCostCenter ||
      !item.EmployeeCostCenterName ||
      !item.Amount ||
      !item.ExpenseMonth ||
      !item.Remarks
    ) {
      alert("All fields are mandatory");
      return;
    }

    if (isNaN(item.Amount)) {
      alert("Amount must be numeric");
      return;
    }

    await sp.web.lists.getByTitle("AccrualSheetList").items.add({
      Username: item.Username,
      Department: item.Department,
      VendorName: item.VendorName,
      VendorCode: item.VendorCode,
      PONumber: item.PONumber,
      GLCode: item.GLCode,
      GLDescription: item.GLDescription,
      EmployeeCostCenter: item.EmployeeCostCenter,
      EmployeeCostCenterName: item.EmployeeCostCenterName,
      Amount: item.Amount,
      ExpenseMonth: item.ExpenseMonth,
      Remarks: item.Remarks,
      DeleteFlag: false,
      Status: "Pending",
    });

    alert("Record Saved Successfully");

    void searchData();
  };

  const Reset = () => {
    setUserName("");
    setDepartment("");
    setVendorName("");
    setPoNumber("");
    setFromDate("");
    setToDate("");

    setData([]);
    setIsSearched(false);
  };

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };
  // const sortedData = [...filteredData].sort((a, b) => b.Id - a.Id);

  const paginatedData = data.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage,
  );

  React.useEffect(() => {
    void checkUserAccess();
    void loadDropdownData();
    void getClosedMonths(); // 👈 ADD THIS
  }, []);

  
  return (
    <div>
      <div className="header">
        <h1>Adjustment Report</h1>
      </div>
      <div className="PaddAll">
        <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
          <select
            value={userName}
            className="form-control"
            onChange={(e) => setUserName(e.target.value)}
          >
            <option value="">All Users</option>
            {userOptions.map((u, i) => (
              <option key={i} value={u}>
                {u}
              </option>
            ))}
          </select>

          <select
            value={department}
            className="form-control"
            onChange={(e) => setDepartment(e.target.value)}
          >
            <option value="">All Departments</option>
            {deptOptions.map((d, i) => (
              <option key={i} value={d}>
                {d}
              </option>
            ))}
          </select>

          <select
            value={vendorName}
            className="form-control"
            onChange={(e) => setVendorName(e.target.value)}
          >
            <option value="">All Vendors</option>
            {vendorOptions.map((v, i) => (
              <option key={i} value={v}>
                {v}
              </option>
            ))}
          </select>

          <select
            value={poNumber}
            className="form-control"
            onChange={(e) => setPoNumber(e.target.value)}
          >
            <option value="">All PO</option>
            {poOptions.map((p, i) => (
              <option key={i} value={p}>
                {p}
              </option>
            ))}
          </select>

          <input
            type="date"
            className="form-control"
            value={fromDate}
            onChange={(e) => setFromDate(e.target.value)}
          />
          <input
            type="date"
            className="form-control"
            value={toDate}
            onChange={(e) => setToDate(e.target.value)}
          />
        </div>
        <div
          className=""
          style={{
            display: "block",
            marginBottom: "20px",
            textAlign: "center",
          }}
        >
          <button onClick={searchData} className="submit-btn">
            Search
          </button>
          <button onClick={exportExcel} className="sendback-btn">
            Export
          </button>
          <button onClick={Reset} className="reset-btn">
            Reset
          </button>
        </div>

        {isSearched && (
          <div className="overflow-x-auto">
            <div className="table-vert-scroll">
              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  style={{ backgroundColor: "#3c3e45" }}
                  className="text-white"
                >
                  <tr>
                    <th>
                      <input
                        type="checkbox"
                        checked={
                          paginatedData.length > 0 &&
                          paginatedData.every((x) => selectedIds.includes(x.Id))
                        }
                        onChange={(e) => {
                          if (e.target.checked) {
                            const ids = paginatedData.map((x) => x.Id);
                            setSelectedIds(ids);
                          } else {
                            setSelectedIds([]);
                          }
                        }}
                      />
                    </th>
                    {/* <th>Select</th> */}
                    <th className="px-4 py-2">Created Date</th>
                    <th className="px-4 py-2">UserName</th>
                    <th className="px-4 py-2">Department</th>
                    <th className="px-4 py-2">Vendor Name</th>
                    <th className="px-4 py-2">Vendor Code</th>
                    <th className="px-4 py-2">PO Number</th>
                    <th className="px-4 py-2">GL Code</th>
                    <th className="px-4 py-2">GL Description</th>
                    <th className="px-4 py-2">Employee Cost Center</th>
                    <th className="px-4 py-2">Employee Cost Center Name</th>
                    <th className="px-4 py-2">Amount</th>
                    <th className="px-4 py-2">Expense Month</th>
                    <th className="px-4 py-2">Remarks</th>
                    {isPerformer && <th>Delete</th>}
                    {isPerformer && <th>Action</th>}
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.length === 0 ? (
                    <tr>
                      <td colSpan={14} style={{ textAlign: "center" }}>
                        No Records Found
                      </td>
                    </tr>
                  ) : (
                    <>
                      {paginatedData.map((item, index) => {
                        const actualIndex =
                          (currentPage - 1) * itemsPerPage + index;

                        return (
                          <tr key={item.Id || actualIndex}>
                            {/* Username */}
                            <td>
                              {item.Id !== 0 && (
                                <input
                                  type="checkbox"
                                  checked={selectedIds.includes(item.Id)}
                                  onChange={(e) => {
                                    if (e.target.checked) {
                                      setSelectedIds([...selectedIds, item.Id]);
                                    } else {
                                      setSelectedIds(
                                        selectedIds.filter(
                                          (id) => id !== item.Id,
                                        ),
                                      );
                                    }
                                  }}
                                />
                              )}
                            </td>
                            {/* <td>{item.Created}</td> */}

                            <td>
                              {item.Created
                                ? new Date(
                                    item.Created.toString(),
                                  ).toLocaleDateString("en-GB")
                                : ""}
                            </td>
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.Username}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "Username",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.Username
                              )}
                            </td>

                            {/* Department */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.Department}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "Department",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.Department
                              )}
                            </td>

                            {/* Vendor */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.VendorName}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "VendorName",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.VendorName
                              )}
                            </td>

                            {/* Vendor Code */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.VendorCode}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "VendorCode",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.VendorCode
                              )}
                            </td>

                            {/* PO */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.PONumber}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "PONumber",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.PONumber
                              )}
                            </td>

                            {/* GL Code */}
                            {/* <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.GLCode}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "GLCode",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.GLCode
                              )}
                            </td> */}
                            <td>
                              <input
                                value={item.GLCode || ""}
                                onChange={(e) =>
                                  updateRow(
                                    actualIndex,
                                    "GLCode",
                                    e.target.value,
                                  )
                                }
                                disabled={
                                  item.Id !== 0 && !editingIds.includes(item.Id)
                                }
                              />
                            </td>

                            {/* GL Desc */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.GLDescription}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "GLDescription",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.GLDescription
                              )}
                            </td>

                            {/* Cost Center */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.EmployeeCostCenter}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "EmployeeCostCenter",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.EmployeeCostCenter
                              )}
                            </td>

                            {/* Cost Center Name */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.EmployeeCostCenterName}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "EmployeeCostCenterName",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.EmployeeCostCenterName
                              )}
                            </td>

                            {/* Amount */}
                            <td>
                              {item.Id === 0 ? (
                                // ✅ Editable for NEW ROW only
                                <input
                                  type="number"
                                  value={item.Amount}
                                  min="0" // ❌ prevents negative from UI
                                  onChange={(e) => {
                                    const val = Number(e.target.value);

                                    // ❌ extra validation (important)
                                    if (val < 0) return;

                                    updateRow(actualIndex, "Amount", val);
                                  }}
                                />
                              ) : (
                                // ✅ Readonly for SEARCH data
                                <input
                                  type="number"
                                  value={item.Amount}
                                  readOnly
                                />
                              )}
                            </td>

                            {/* Month */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.ExpenseMonth}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "ExpenseMonth",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                formatMonth(item.ExpenseMonth)
                              )}
                            </td>

                            {/* Remarks */}
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.Remarks}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "Remarks",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.Remarks
                              )}
                            </td>

                            {/* Delete */}
                            <td>
                              {item.Id === 0 ? (
                                <button
                                  style={{ color: "orange" }}
                                  onClick={() => removeNewRow(actualIndex)}
                                >
                                  Remove
                                </button>
                              ) : (
                                isPerformer && (
                                  <button
                                    style={{ color: "red" }}
                                    onClick={() => deleteItem(item.Id)}
                                  >
                                    Delete
                                  </button>
                                )
                              )}
                            </td>
                            {isPerformer && (
                              <td>
                                {item.Id !== 0 &&
                                  (!editingIds.includes(item.Id) ? (
                                    <button
                                      onClick={() =>
                                        setEditingIds([...editingIds, item.Id])
                                      }
                                    >
                                      Edit
                                    </button>
                                  ) : (
                                    <button onClick={() => updateGLCode(item)}>
                                      Save
                                    </button>
                                  ))}
                              </td>
                            )}
                          </tr>
                        );
                      })}

                      {/* Add Row Buttons */}
                      <tr>
                        <td colSpan={12}></td>
                        {isPerformer && (
                          <>
                            <td>
                              <button
                                className="submit-btn"
                                onClick={addNewRow}
                              >
                                Add New
                              </button>
                            </td>
                            <td>
                              <button onClick={saveAllRows}>Save All</button>
                            </td>
                          </>
                        )}
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
            {isPerformer && (
              <button
                onClick={deleteSelectedItems}
                className="sendback-btn"
                style={{ backgroundColor: "red" }}
              >
                Delete Selected
              </button>
            )}
            {/* Pagination */}
            <div className="flex justify-center mt-6 overflow-x-auto">
              <div
                className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow"
                style={{ textAlign: "end" }}
              >
                {/* Previous Button */}
                <button
                  onClick={() => handlePageChange(currentPage - 1)}
                  disabled={currentPage === 1}
                  style={{
                    backgroundColor: "#fff",
                    border: "1px solid #000 !important",
                    marginRight: "5px",
                    opacity: currentPage === 1 ? 0.5 : 1,
                  }}
                  className="px-3 py-1 border rounded"
                >
                  <img src={Left} alt="" width={15} />
                </button>
                {/* Main Page Numbers */}
                {Array.from({ length: totalPages }, (_, i) => i + 1)
                  .filter((page) => Math.abs(page - currentPage) <= 2)
                  .map((page) => (
                    <button
                      key={page}
                      onClick={() => handlePageChange(page)}
                      style={{
                        backgroundColor:
                          currentPage === page ? "#3c3e45" : "#fff",
                        color: currentPage === page ? "#fff" : "#000",
                        fontWeight: currentPage === page ? "bold" : "normal",
                        margin: currentPage === page ? "5px" : "5px",
                      }}
                      className="px-3 py-1 border rounded"
                    >
                      {page}
                    </button>
                  ))}

                {/* Next Button */}
                <button
                  onClick={() => handlePageChange(currentPage + 1)}
                  disabled={currentPage === totalPages}
                  style={{
                    backgroundColor: "#fff",
                    border: "1px solid #000 !important",
                    marginLeft: "5px",
                    opacity: currentPage === totalPages ? 0.5 : 1,
                  }}
                  className="px-3 py-1 border rounded"
                >
                  <img src={Right} alt="" width={15} />
                </button>
              </div>
            </div>
          </div>
        )}
        {isPerformer && (
          <div>
            <div className="row">
              <div
                className="col-md-12 col-sm-12"
                style={{
                  display: "flex",
                  gap: "15px",
                  margin: "0px 10px",
                  alignItems: "center",
                }}
              >
                {/* <div>
                  <label>Select Excel File</label>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      if (e.target.files && e.target.files.length > 0) {
                        setFile(e.target.files[0]);
                      }
                    }}
                  />
                </div> */}
                <div>
                  <button
                    className="primaryBtn"
                    style={{ margin: "0px" }}
                    onClick={freezeData}
                  >
                    Freeze send to GL
                  </button>
                </div>
                <div>
                  <button onClick={handleExit} className="Reject-btn">
                    {" "}
                    Exit{" "}
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
