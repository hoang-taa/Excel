import { useState, useRef } from "react";
import * as XLSX from "xlsx/xlsx.mjs";
import { saveAs } from "file-saver";
import FileProgressUploading from "./FileProgressUploading";
import TableData from "./Table";

function App() {
  const [file1Data, setFile1Data] = useState([]);
  const [file2Data, setFile2Data] = useState([]);
  const [file3Data, setFile3Data] = useState([]);
  const [fileSitesCheck, setFileSitesCheck] = useState([]);

  const [statusFileCellConfig, setStatusFileCellConfig] = useState({
    fileName: "",
    fileSize: "",
    progressUpload: 0,
  });
  const [statusFileInventory, setStatusFileInventory] = useState({
    fileName: "",
    fileSize: "",
    progressUpload: 0,
  });
  const [statusFileCapacity, setStatusFileCapacity] = useState({
    fileName: "",
    fileSize: "",
    progressUpload: 0,
  });

  const [statusFileSitesCheck, setStatusFileSitesCheck] = useState({
    fileName: "",
    fileSize: "",
    progressUpload: 0,
  });

  const [uniqueSiteNames, setUniqueSiteNames] = useState([]);
  const [errorAlert, setErrorAlert] = useState("");

  const handleReadFileCellConfig = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onprogress = function (event) {
      if (event.lengthComputable) {
        const percentLoaded = Math.round((event.loaded / event.total) * 100);
        if (file) {
          setStatusFileCellConfig({
            fileName: file?.name,
            fileSize: formatBytes(file?.size),
            progressUpload: percentLoaded,
          });
        }
      }
    };
    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]]
      );

      const objectsList = [];
      jsonData.forEach(function (row) {
        var obj = {};
        Object.keys(row).forEach(function (key) {
          obj[key.trim()] = row[key];
        });
        objectsList.push(obj);
      });

      function getUniqueSiteNames(data) {
        const uniqueSiteNames = {};
        data.forEach((item) => {
          const sitename = item["SITENAME"];
          if (!uniqueSiteNames[sitename]) {
            uniqueSiteNames[sitename] = sitename.slice(-11);
          }
        });
        return Object.values(uniqueSiteNames);
      }

      const uniqueSiteNames = getUniqueSiteNames(objectsList);
      setUniqueSiteNames(uniqueSiteNames);

      setFile1Data(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  function formatBytes(bytes) {
    var marker = 1024; // Change to 1000 if required
    var decimal = 3; // Change as required
    var kiloBytes = marker; // One Kilobyte is 1024 bytes
    var megaBytes = marker * marker; // One MB is 1024 KB
    var gigaBytes = marker * marker * marker; // One GB is 1024 MB

    // return bytes if less than a KB
    if (bytes < kiloBytes) return bytes + " Bytes";
    // return KB if less than a MB
    else if (bytes < megaBytes)
      return (bytes / kiloBytes).toFixed(decimal) + " KB";
    // return MB if less than a GB
    else if (bytes < gigaBytes)
      return (bytes / megaBytes).toFixed(decimal) + " MB";
    // return GB if less than a TB
    else return (bytes / gigaBytes).toFixed(decimal) + " GB";
  }

  const handleReadFileInventory = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onprogress = function (event) {
      if (event.lengthComputable) {
        const percentLoaded = Math.round((event.loaded / event.total) * 100);
        if (file) {
          setStatusFileInventory({
            fileName: file?.name,
            fileSize: formatBytes(file?.size),
            progressUpload: percentLoaded,
          });
        }
      }
    };
    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]]
      );
      setFile2Data(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleReadFileCapacity = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onprogress = function (event) {
      if (event.lengthComputable) {
        const percentLoaded = Math.round((event.loaded / event.total) * 100);
        if (file) {
          setStatusFileCapacity({
            fileName: file?.name,
            fileSize: formatBytes(file?.size),
            progressUpload: percentLoaded,
          });
        }
      }
    };

    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]]
      );
      setFile3Data(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleReadFileSitesCheck = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onprogress = function (event) {
      if (event.lengthComputable) {
        const percentLoaded = Math.round((event.loaded / event.total) * 100);
        if (file) {
          setStatusFileSitesCheck({
            fileName: file?.name,
            fileSize: formatBytes(file?.size),
            progressUpload: percentLoaded,
          });
        }
      }
    };
    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]]
      );
      setFileSitesCheck(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const inputFileCellConfig = useRef(null);
  const inputFileInventory = useRef(null);
  const inputFileCapacity = useRef(null);
  const inputFileSitesCheck = useRef(null);
  const [finalDataTable, setFinalDataTable] = useState([]);
  const clearData = () => {
    if (inputFileCellConfig.current) {
      inputFileCellConfig.current.value = "";
      inputFileCellConfig.current.type = "text";
      inputFileCellConfig.current.type = "file";
    }

    if (inputFileInventory.current) {
      inputFileInventory.current.value = "";
      inputFileInventory.current.type = "text";
      inputFileInventory.current.type = "file";
    }

    if (inputFileCapacity.current) {
      inputFileCapacity.current.value = "";
      inputFileCapacity.current.type = "text";
      inputFileCapacity.current.type = "file";
    }

    if (inputFileSitesCheck.current) {
      inputFileSitesCheck.current.value = "";
      inputFileSitesCheck.current.type = "text";
      inputFileSitesCheck.current.type = "file";
    }

    setStatusFileCellConfig({
      fileName: "",
      fileSize: "",
      progressUpload: 0,
    });

    setStatusFileInventory({
      fileName: "",
      fileSize: "",
      progressUpload: 0,
    });

    setStatusFileCapacity({
      fileName: "",
      fileSize: "",
      progressUpload: 0,
    });

    setStatusFileSitesCheck({
      fileName: "",
      fileSize: "",
      progressUpload: 0,
    });

    setFile1Data([]);
    setFile2Data([]);
    setFile3Data([]);
    setFileSitesCheck([]);
    setUniqueSiteNames([]);
    setErrorAlert("");
    setFinalDataTable([]);
  };

  const mergeData = () => {
    if (file1Data.length === 0) {
      setErrorAlert("Please import file Cell Config");
    } else if (file2Data.length === 0) {
      setErrorAlert("Please import file Inventory");
    } else if (file3Data.length === 0) {
      setErrorAlert("Please import file Capacity");
    } else {
      setErrorAlert("");
    }

    const getExistData = result.filter(
      (item) => item.QuantityTypeUBBP && item.QuantityTypeRRU && item.CellCounts
    );

    const finalData = calculateMaxCellMoran(getExistData ?? []).map((item) => {
      return {
        "Tên site": item.NEName,
        "Số lượng & chủng loại UBBP": item.QuantityTypeUBBP,
        "Số lượng và chủng loại RRU": item.QuantityTypeRRU,
        "Số cell hiện tại": item.CellCounts,
        "Số cell hỗ trợ tối đa": item.MaxCell,
        "Đáp ứng Moran (Y/N)": item.ResponseMoran,
      };
    });

    setFinalDataTable(finalData);
    exportToExcel(finalData);
  };

  const filterSitesList = () => {
    if (file1Data.length === 0) {
      setErrorAlert("Please import file Cell Config");
    } else if (file2Data.length === 0) {
      setErrorAlert("Please import file Inventory");
    } else if (file3Data.length === 0) {
      setErrorAlert("Please import file Capacity");
    } else {
      setErrorAlert("");
    }

    const getExistData = result.filter(
      (item) => item.QuantityTypeUBBP && item.QuantityTypeRRU && item.CellCounts
    );

    const listFilter = calculateMaxCellMoran(getExistData ?? [])
      .filter((obj1) =>
        fileSitesCheck.some((obj2) => obj1.NEName === obj2.NEName)
      )
      .map((item) => {
        return {
          "Tên site": item.NEName,
          "Số lượng & chủng loại UBBP": item.QuantityTypeUBBP,
          "Số lượng và chủng loại RRU": item.QuantityTypeRRU,
          "Số cell hiện tại": item.CellCounts,
          "Số cell hỗ trợ tối đa": item.MaxCell,
          "Đáp ứng Moran (Y/N)": item.ResponseMoran,
        };
      });
    setFinalDataTable(listFilter);

    exportToExcel(listFilter);
  };

  const calculateMaxCellMoran = (data) => {
    // Tạo một đối tượng map chứa thông tin về Max cell của từng loại UBBP
    const ubbpMaxMap = {};
    file3Data.forEach((item) => {
      ubbpMaxMap[item["UBBP Type"]] = item["Max cell"];
    });

    // Duyệt qua mảng 1 và cập nhật giá trị MaxCell và ResponseMoran tương ứng
    const modifiedArray = data.map((item) => {
      // Tách và lấy thông tin về loại UBBP và số lượng từ chuỗi
      const ubbpInfo = item["QuantityTypeUBBP"].split(" + ");
      let totalMaxCell = 0;
      ubbpInfo.forEach((info) => {
        const [quantity, ubbpType] = info.split(" ");
        const maxCell = ubbpMaxMap[ubbpType]; // Lấy giá trị Max cell từ map
        totalMaxCell += parseInt(quantity) * maxCell; // Tính tổng Max cell
      });
      const maxCell = totalMaxCell;
      const cellCounts = item["CellCounts"].split("+").reduce((acc, curr) => {
        const count = parseInt(curr.trim().split(" ")[0]);
        return acc + count;
      }, 0);
      const responseMoran = maxCell - cellCounts;
      return {
        ...item,
        MaxCell: maxCell,
        ResponseMoran: responseMoran > 0 ? "Y" : "N",
      }; // Thêm thuộc tính MaxCell và ResponseMoran vào đối tượng
    });
    return modifiedArray;
  };

  const exportToExcel = (data) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();

    var wscols = [
      { wch: 20 },
      { wch: 35 },
      { wch: 35 },
      { wch: 35 },
      { wch: 20 },
      { wch: 20 },
    ];

    ws["!cols"] = wscols;
    XLSX.utils.book_append_sheet(wb, ws, "List Data");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });

    function s2ab(s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    }

    saveAs(
      new Blob([s2ab(wbout)], { type: "application/octet-stream" }),
      "outputExport.xlsx"
    );
  };

  const UBBPCounts = file2Data.reduce((acc, obj) => {
    // Nếu trường "Board Type" chứa chuỗi "UBBP", tăng biến đếm UBBP lên
    if (
      typeof obj["Board Type"] == "string" &&
      obj["Board Type"].includes("UBBP")
    ) {
      const UBBPType = obj["Board Type"].split("UBBP")[1]; // Lấy phần loại UBBP sau chuỗi "UBBP"
      acc[obj.NEName] = acc[obj.NEName] || {};
      acc[obj.NEName][UBBPType] = (acc[obj.NEName][UBBPType] || 0) + 1;
    }

    return acc;
  }, {});

  const RRUCounts = file2Data.reduce((acc, obj) => {
    // Nếu trường "Board Type" chứa chuỗi "UBBP", tăng biến đếm UBBP lên
    if (
      typeof obj["Manufacturer Data"] == "string" &&
      obj["Manufacturer Data"].includes("RRU")
    ) {
      const RRUType = obj["Manufacturer Data"].split(",")[0]; // Lấy phần loại UBBP sau chuỗi "UBBP"
      acc[obj.NEName] = acc[obj.NEName] || {};
      acc[obj.NEName][RRUType] = (acc[obj.NEName][RRUType] || 0) + 1;
    }
    return acc;
  }, {});

  const CellCounts = file1Data.reduce((acc, obj) => {
    const MIMOType = obj["MIMO"];
    const siteName = obj["SITENAME"].slice(-11);
    acc[siteName] = acc[siteName] || {}; // Tạo một đối tượng con nếu chưa tồn tại
    acc[siteName][MIMOType] = (acc[siteName][MIMOType] || 0) + 1; // Tăng biến đếm cho loại MIMO tương ứng
    return acc;
  }, {});

  // Tạo mảng kết quả từ Array1 và UBBPCounts
  const result = uniqueSiteNames.map((NEName) => {
    return {
      NEName,
      QuantityTypeUBBP: Object.entries(UBBPCounts[NEName] || {})
        .map(([UBBPType, count]) => `${count} UBBP${UBBPType}`)
        .join(" + "),
      QuantityTypeRRU: Object.entries(RRUCounts[NEName] || {})
        .map(([UBBPType, count]) => `${count} ${UBBPType}`)
        .join(" + "),
      CellCounts: Object.entries(CellCounts[NEName] || {})
        .map(([cellNumber, count]) => {
          return `${count} cells ${cellNumber}`;
        })
        .join(" + "),
    };
  });
  return (
    <>
      <body className='bg-gray-50 dark:bg-slate-900'>
        <header className='sticky top-0 inset-x-0 flex flex-wrap sm:justify-start sm:flex-nowrap z-[48] w-full bg-white border-b text-sm py-2.5 sm:py-4 lg:ps-64 dark:bg-gray-800 dark:border-gray-700'>
          <nav
            className='flex basis-full items-center w-full mx-auto px-4 sm:px-6 md:px-8'
            aria-label='Global'
          >
            <div className='me-5 lg:me-0 lg:hidden'>
              <a
                className='flex-none text-xl font-semibold dark:text-white'
                href='#'
                aria-label='Tool Excel'
              >
                <a
                  target='_blank'
                  href='https://icons8.com/icon/117561/microsoft-excel-2019'
                >
                  Excel
                </a>{" "}
                icon by{" "}
                <a target='_blank' href='https://icons8.com'>
                  Icons8
                </a>{" "}
                Tool Excel
              </a>
            </div>

            <div className='w-full flex items-center justify-end ms-auto sm:justify-between sm:gap-x-3 sm:order-3'>
              <div className='sm:hidden'>
                <button
                  type='button'
                  className='w-[2.375rem] h-[2.375rem] inline-flex justify-center items-center gap-x-2 text-sm font-semibold rounded-full border border-transparent text-gray-800 hover:bg-gray-100 disabled:opacity-50 disabled:pointer-events-none dark:text-white dark:hover:bg-gray-700 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                >
                  <svg
                    className='flex-shrink-0 size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <circle cx='11' cy='11' r='8' />
                    <path d='m21 21-4.3-4.3' />
                  </svg>
                </button>
              </div>

              <div className='hidden sm:block'>
                <label htmlFor='icon' className='sr-only'>
                  Search
                </label>
                <div className='relative'>
                  <div className='absolute inset-y-0 start-0 flex items-center pointer-events-none z-20 ps-4'>
                    <svg
                      className='flex-shrink-0 size-4 text-gray-400'
                      xmlns='http://www.w3.org/2000/svg'
                      width='24'
                      height='24'
                      viewBox='0 0 24 24'
                      fill='none'
                      stroke='currentColor'
                      strokeWidth='2'
                      strokeLinecap='round'
                      strokeLinejoin='round'
                    >
                      <circle cx='11' cy='11' r='8' />
                      <path d='m21 21-4.3-4.3' />
                    </svg>
                  </div>
                  <input
                    type='text'
                    id='icon'
                    name='icon'
                    className='py-2 px-4 ps-11 block w-full border-gray-200 rounded-lg text-sm focus:border-blue-500 focus:ring-blue-500 disabled:opacity-50 disabled:pointer-events-none dark:bg-slate-900 dark:border-gray-700 dark:text-gray-400 dark:focus:ring-gray-600'
                    placeholder='Search'
                  />
                </div>
              </div>

              <div className='flex flex-row items-center justify-end gap-2'>
                <button
                  type='button'
                  className='w-[2.375rem] h-[2.375rem] inline-flex justify-center items-center gap-x-2 text-sm font-semibold rounded-full border border-transparent text-gray-800 hover:bg-gray-100 disabled:opacity-50 disabled:pointer-events-none dark:text-white dark:hover:bg-gray-700 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                >
                  <svg
                    className='flex-shrink-0 size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9' />
                    <path d='M10.3 21a1.94 1.94 0 0 0 3.4 0' />
                  </svg>
                </button>
                <button
                  type='button'
                  className='w-[2.375rem] h-[2.375rem] inline-flex justify-center items-center gap-x-2 text-sm font-semibold rounded-full border border-transparent text-gray-800 hover:bg-gray-100 disabled:opacity-50 disabled:pointer-events-none dark:text-white dark:hover:bg-gray-700 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                  data-hs-offcanvas='#hs-offcanvas-right'
                >
                  <svg
                    className='flex-shrink-0 size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='M22 12h-4l-3 9L9 3l-3 9H2' />
                  </svg>
                </button>

                <div className='hs-dropdown relative inline-flex [--placement:bottom-right]'>
                  <button
                    id='hs-dropdown-with-header'
                    type='button'
                    className='w-[2.375rem] h-[2.375rem] inline-flex justify-center items-center gap-x-2 text-sm font-semibold rounded-full border border-transparent text-gray-800 hover:bg-gray-100 disabled:opacity-50 disabled:pointer-events-none dark:text-white dark:hover:bg-gray-700 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                  >
                    <img
                      className='inline-block size-[38px] rounded-full ring-2 ring-white dark:ring-gray-800'
                      src='https://images.unsplash.com/photo-1568602471122-7832951cc4c5?ixlib=rb-4.0.3&ixid=MnwxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8&auto=format&fit=facearea&facepad=2&w=320&h=320&q=80'
                      alt='Image Description'
                    />
                  </button>

                  <div
                    className='hs-dropdown-menu transition-[opacity,margin] duration hs-dropdown-open:opacity-100 opacity-0 hidden min-w-60 bg-white shadow-md rounded-lg p-2 dark:bg-gray-800 dark:border dark:border-gray-700'
                    aria-labelledby='hs-dropdown-with-header'
                  >
                    <div className='py-3 px-5 -m-2 bg-gray-100 rounded-t-lg dark:bg-gray-700'>
                      <p className='text-sm text-gray-500 dark:text-gray-400'>
                        Signed in as
                      </p>
                      <p className='text-sm font-medium text-gray-800 dark:text-gray-300'>
                        Tran-Sy
                      </p>
                    </div>
                    <div className='mt-2 py-2 first:pt-0 last:pb-0'>
                      <a
                        className='flex items-center gap-x-3.5 py-2 px-3 rounded-lg text-sm text-gray-800 hover:bg-gray-100 focus:ring-2 focus:ring-blue-500 dark:text-gray-400 dark:hover:bg-gray-700 dark:hover:text-gray-300'
                        href='#'
                      >
                        <svg
                          className='flex-shrink-0 size-4'
                          xmlns='http://www.w3.org/2000/svg'
                          width='24'
                          height='24'
                          viewBox='0 0 24 24'
                          fill='none'
                          stroke='currentColor'
                          strokeWidth='2'
                          strokeLinecap='round'
                          strokeLinejoin='round'
                        >
                          <path d='M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2' />
                          <circle cx='9' cy='7' r='4' />
                          <path d='M22 21v-2a4 4 0 0 0-3-3.87' />
                          <path d='M16 3.13a4 4 0 0 1 0 7.75' />
                        </svg>
                        Logout
                      </a>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </nav>
        </header>

        <div className='sticky top-0 inset-x-0 z-20 bg-white border-y px-4 sm:px-6 md:px-8 lg:hidden dark:bg-gray-800 dark:border-gray-700'>
          <div className='flex items-center py-4'>
            <button
              type='button'
              className='text-gray-500 hover:text-gray-600'
              data-hs-overlay='#application-sidebar'
              aria-controls='application-sidebar'
              aria-label='Toggle navigation'
            >
              <span className='sr-only'>Toggle Navigation</span>
              <svg
                className='flex-shrink-0 size-4'
                xmlns='http://www.w3.org/2000/svg'
                width='24'
                height='24'
                viewBox='0 0 24 24'
                fill='none'
                stroke='currentColor'
                strokeWidth='2'
                strokeLinecap='round'
                strokeLinejoin='round'
              >
                <line x1='3' x2='21' y1='6' y2='6' />
                <line x1='3' x2='21' y1='12' y2='12' />
                <line x1='3' x2='21' y1='18' y2='18' />
              </svg>
            </button>

            <ol
              className='ms-3 flex items-center whitespace-nowrap'
              aria-label='Breadcrumb'
            >
              <li className='flex items-center text-sm text-gray-800 dark:text-gray-400'>
                Application Layout
                <svg
                  className='flex-shrink-0 mx-3 overflow-visible size-2.5 text-gray-400 dark:text-gray-600'
                  width='16'
                  height='16'
                  viewBox='0 0 16 16'
                  fill='none'
                  xmlns='http://www.w3.org/2000/svg'
                >
                  <path
                    d='M5 1L10.6869 7.16086C10.8637 7.35239 10.8637 7.64761 10.6869 7.83914L5 14'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                  />
                </svg>
              </li>
              <li
                className='text-sm font-semibold text-gray-800 truncate dark:text-gray-400'
                aria-current='page'
              >
                Tool Excel
              </li>
            </ol>
          </div>
        </div>

        <div
          id='application-sidebar'
          className='float-left hs-overlay hs-overlay-open:translate-x-0 -translate-x-full transition-all duration-300 transform hidden fixed top-0 start-0 bottom-0 z-[60] w-64 bg-white border-e border-gray-200 pt-7 pb-10 overflow-y-auto lg:block lg:translate-x-0 lg:end-auto lg:bottom-0 [&::-webkit-scrollbar]:w-2 [&::-webkit-scrollbar-thumb]:rounded-full [&::-webkit-scrollbar-track]:bg-gray-100 [&::-webkit-scrollbar-thumb]:bg-gray-300 dark:[&::-webkit-scrollbar-track]:bg-slate-700 dark:[&::-webkit-scrollbar-thumb]:bg-slate-500 dark:bg-gray-800 dark:border-gray-700'
        >
          <div className='px-6'>
            <a
              className='flex-none text-xl font-semibold dark:text-white dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
              href='#'
              aria-label='Tool Excel'
            >
              <div className='flex items-center'>
                <svg
                  xmlns='http://www.w3.org/2000/svg'
                  x='0px'
                  y='0px'
                  width='30'
                  height='30'
                  viewBox='0 0 48 48'
                >
                  <path
                    fill='#169154'
                    d='M29,6H15.744C14.781,6,14,6.781,14,7.744v7.259h15V6z'
                  ></path>
                  <path
                    fill='#18482a'
                    d='M14,33.054v7.202C14,41.219,14.781,42,15.743,42H29v-8.946H14z'
                  ></path>
                  <path
                    fill='#0c8045'
                    d='M14 15.003H29V24.005000000000003H14z'
                  ></path>
                  <path fill='#17472a' d='M14 24.005H29V33.055H14z'></path>
                  <g>
                    <path
                      fill='#29c27f'
                      d='M42.256,6H29v9.003h15V7.744C44,6.781,43.219,6,42.256,6z'
                    ></path>
                    <path
                      fill='#27663f'
                      d='M29,33.054V42h13.257C43.219,42,44,41.219,44,40.257v-7.202H29z'
                    ></path>
                    <path
                      fill='#19ac65'
                      d='M29 15.003H44V24.005000000000003H29z'
                    ></path>
                    <path fill='#129652' d='M29 24.005H44V33.055H29z'></path>
                  </g>
                  <path
                    fill='#0c7238'
                    d='M22.319,34H5.681C4.753,34,4,33.247,4,32.319V15.681C4,14.753,4.753,14,5.681,14h16.638 C23.247,14,24,14.753,24,15.681v16.638C24,33.247,23.247,34,22.319,34z'
                  ></path>
                  <path
                    fill='#fff'
                    d='M9.807 19L12.193 19 14.129 22.754 16.175 19 18.404 19 15.333 24 18.474 29 16.123 29 14.013 25.07 11.912 29 9.526 29 12.719 23.982z'
                  ></path>
                </svg>
                <span className='ml-2 font-normal text-teal-500'>
                  Tool Excel
                </span>
              </div>
            </a>
          </div>

          <nav
            className='hs-accordion-group p-6 w-full flex flex-col flex-wrap'
            data-hs-accordion-always-open
          >
            <ul className='space-y-1.5'>
              <li>
                <a
                  className='flex items-center gap-x-3.5 py-2 px-2.5 bg-gray-100 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-900 dark:text-white dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                  href='#'
                >
                  <svg
                    className='flex-shrink-0 size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z' />
                    <polyline points='9 22 9 12 15 12 15 22' />
                  </svg>
                  Tool Excel
                </a>
              </li>

              <li className='hs-accordion' id='users-accordion'>
                <button
                  type='button'
                  className='hs-accordion-toggle w-full text-start flex items-center gap-x-3.5 py-2 px-2.5 hs-accordion-active:text-teal-500 hs-accordion-active:hover:bg-transparent text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:hover:bg-gray-900 dark:text-slate-400 dark:hover:text-slate-300 dark:hs-accordion-active:text-white dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                >
                  <svg
                    className='flex-shrink-0 size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2' />
                    <circle cx='9' cy='7' r='4' />
                    <path d='M22 21v-2a4 4 0 0 0-3-3.87' />
                    <path d='M16 3.13a4 4 0 0 1 0 7.75' />
                  </svg>
                  Users
                  <svg
                    className='hs-accordion-active:block ms-auto hidden size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='m18 15-6-6-6 6' />
                  </svg>
                  <svg
                    className='hs-accordion-active:hidden ms-auto block size-4'
                    xmlns='http://www.w3.org/2000/svg'
                    width='24'
                    height='24'
                    viewBox='0 0 24 24'
                    fill='none'
                    stroke='currentColor'
                    strokeWidth='2'
                    strokeLinecap='round'
                    strokeLinejoin='round'
                  >
                    <path d='m6 9 6 6 6-6' />
                  </svg>
                </button>

                <div
                  id='users-accordion-child'
                  className='hs-accordion-content w-full overflow-hidden transition-[height] duration-300 hidden'
                >
                  <ul
                    className='hs-accordion-group ps-3 pt-2'
                    data-hs-accordion-always-open
                  >
                    <li className='hs-accordion' id='users-accordion-sub-1'>
                      <button
                        type='button'
                        className='hs-accordion-toggle w-full text-start flex items-center gap-x-3.5 py-2 px-2.5 hs-accordion-active:text-teal-500 hs-accordion-active:hover:bg-transparent text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:hover:bg-gray-900 dark:text-slate-400 dark:hover:text-slate-300 dark:hs-accordion-active:text-white dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                      >
                        Sub Menu 1
                        <svg
                          className='hs-accordion-active:block ms-auto hidden size-4'
                          xmlns='http://www.w3.org/2000/svg'
                          width='24'
                          height='24'
                          viewBox='0 0 24 24'
                          fill='none'
                          stroke='currentColor'
                          strokeWidth='2'
                          strokeLinecap='round'
                          strokeLinejoin='round'
                        >
                          <path d='m18 15-6-6-6 6' />
                        </svg>
                        <svg
                          className='hs-accordion-active:hidden ms-auto block size-4'
                          xmlns='http://www.w3.org/2000/svg'
                          width='24'
                          height='24'
                          viewBox='0 0 24 24'
                          fill='none'
                          stroke='currentColor'
                          strokeWidth='2'
                          strokeLinecap='round'
                          strokeLinejoin='round'
                        >
                          <path d='m6 9 6 6 6-6' />
                        </svg>
                      </button>

                      <div
                        id='users-accordion-sub-1-child'
                        className='hs-accordion-content w-full overflow-hidden transition-[height] duration-300 hidden'
                      >
                        <ul className='pt-2 ps-2'>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 1
                            </a>
                          </li>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 2
                            </a>
                          </li>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 3
                            </a>
                          </li>
                        </ul>
                      </div>
                    </li>
                    <li className='hs-accordion' id='users-accordion-sub-2'>
                      <button
                        type='button'
                        className='hs-accordion-toggle w-full text-start flex items-center gap-x-3.5 py-2 px-2.5 hs-accordion-active:text-teal-500 hs-accordion-active:hover:bg-transparent text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:hover:bg-gray-900 dark:text-slate-400 dark:hover:text-slate-300 dark:hs-accordion-active:text-white dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                      >
                        Sub Menu 2
                        <svg
                          className='hs-accordion-active:block ms-auto hidden size-4'
                          xmlns='http://www.w3.org/2000/svg'
                          width='24'
                          height='24'
                          viewBox='0 0 24 24'
                          fill='none'
                          stroke='currentColor'
                          strokeWidth='2'
                          strokeLinecap='round'
                          strokeLinejoin='round'
                        >
                          <path d='m18 15-6-6-6 6' />
                        </svg>
                        <svg
                          className='hs-accordion-active:hidden ms-auto block size-4'
                          xmlns='http://www.w3.org/2000/svg'
                          width='24'
                          height='24'
                          viewBox='0 0 24 24'
                          fill='none'
                          stroke='currentColor'
                          strokeWidth='2'
                          strokeLinecap='round'
                          strokeLinejoin='round'
                        >
                          <path d='m6 9 6 6 6-6' />
                        </svg>
                      </button>

                      <div
                        id='users-accordion-sub-2-child'
                        className='hs-accordion-content w-full overflow-hidden transition-[height] duration-300 hidden ps-2'
                      >
                        <ul className='pt-2 ps-2'>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 1
                            </a>
                          </li>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 2
                            </a>
                          </li>
                          <li>
                            <a
                              className='flex items-center gap-x-3.5 py-2 px-2.5 text-sm text-slate-700 rounded-lg hover:bg-gray-100 dark:bg-gray-800 dark:text-slate-400 dark:hover:text-slate-300 dark:focus:outline-none dark:focus:ring-1 dark:focus:ring-gray-600'
                              href='#'
                            >
                              Link 3
                            </a>
                          </li>
                        </ul>
                      </div>
                    </li>
                  </ul>
                </div>
              </li>
            </ul>
          </nav>
        </div>

        <div className='float-right w-[calc(100%_-_16rem)] flex h-lvh  items-start flex-col'>
          <div className='font-[sans-serif] w-full mt-2 flex flex-row'>
            <div className='font-[sans-serif] mx-auto min-w-60'>
              {statusFileCellConfig.progressUpload !== 100 ? (
                <>
                  <label className='text-base font-semibold text-teal-500 m-2 block'>
                    Upload file Cell Config
                  </label>
                  <input
                    type='file'
                    onChange={handleReadFileCellConfig}
                    ref={inputFileCellConfig}
                    accept='.xlsx'
                    className='w-full text-black text-sm bg-white border file:cursor-pointer cursor-pointer file:border-0 file:py-2.5 file:px-4 file:bg-gray-100 file:hover:bg-gray-200 file:text-black rounded'
                  />
                </>
              ) : (
                <FileProgressUploading
                  fileName={statusFileCellConfig.fileName}
                  fileSize={statusFileCellConfig.fileSize}
                  progressUpload={statusFileCellConfig.progressUpload}
                />
              )}
            </div>

            <div className='font-[sans-serif] mx-auto'>
              {statusFileInventory.progressUpload !== 100 ? (
                <>
                  <label className='text-base font-semibold text-teal-500 m-2 block'>
                    Upload file Inventory
                  </label>
                  <input
                    type='file'
                    ref={inputFileInventory}
                    onChange={handleReadFileInventory}
                    accept='.xlsx, .csv'
                    className='w-full text-black text-sm bg-white border file:cursor-pointer cursor-pointer file:border-0 file:py-2.5 file:px-4 file:bg-gray-100 file:hover:bg-gray-200 file:text-black rounded'
                  />
                </>
              ) : (
                <FileProgressUploading
                  fileName={statusFileInventory.fileName}
                  fileSize={statusFileInventory.fileSize}
                  progressUpload={statusFileInventory.progressUpload}
                />
              )}
            </div>

            <div className='font-[sans-serif] mx-auto'>
              {statusFileCapacity.progressUpload !== 100 ? (
                <>
                  <label className='text-base font-semibold text-teal-500 m-2 block'>
                    Upload file Capacity
                  </label>
                  <input
                    type='file'
                    ref={inputFileCapacity}
                    onChange={handleReadFileCapacity}
                    accept='.xlsx, .csv'
                    className='w-full text-black text-sm bg-white border file:cursor-pointer cursor-pointer file:border-0 file:py-2.5 file:px-4 file:bg-gray-100 file:hover:bg-gray-200 file:text-black rounded'
                  />
                </>
              ) : (
                <FileProgressUploading
                  fileName={statusFileCapacity.fileName}
                  fileSize={statusFileCapacity.fileSize}
                  progressUpload={statusFileCapacity.progressUpload}
                />
              )}
            </div>

            <div className='font-[sans-serif] mx-auto'>
              {statusFileSitesCheck.progressUpload !== 100 ? (
                <>
                  <label className='text-base font-semibold text-teal-500 m-2 block'>
                    Upload file Sites List Check
                  </label>
                  <input
                    type='file'
                    ref={inputFileSitesCheck}
                    onChange={handleReadFileSitesCheck}
                    accept='.xlsx, .csv'
                    className='w-full text-black text-sm bg-white border file:cursor-pointer cursor-pointer file:border-0 file:py-2.5 file:px-4 file:bg-gray-100 file:hover:bg-gray-200 file:text-black rounded'
                  />
                </>
              ) : (
                <FileProgressUploading
                  fileName={statusFileSitesCheck.fileName}
                  fileSize={statusFileSitesCheck.fileSize}
                  progressUpload={statusFileSitesCheck.progressUpload}
                />
              )}
            </div>
          </div>

          <div className='font-[sans-serif] w-full mt-2 flex flex-col'>
            <div className='font-[sans-serif] space-x-4 space-y-4 text-center'>
              <button
                type='button'
                onClick={mergeData}
                className='px-6 py-2 rounded text-white text-sm tracking-wider font-medium outline-none border-2 border-teal-500 bg-teal-500 hover:bg-transparent hover:text-black transition-all duration-300 relative active:top-[1px]'
              >
                Export Excel
              </button>

              <button
                type='button'
                onClick={clearData}
                className='px-6 py-2 rounded text-black text-sm tracking-wider font-medium outline-none border-2 border-teal-500 relative active:top-[1px]'
              >
                Clear Data
              </button>

              <button
                type='button'
                onClick={filterSitesList}
                className='px-6 py-2 rounded text-black text-sm tracking-wider font-medium outline-none border-2 border-teal-500 relative active:top-[1px]'
              >
                Filter by sites list
              </button>
            </div>
            {errorAlert && (
              <div
                className='bg-red-100 text-red-800 pl-4 pr-10 py-4 rounded relative'
                role='alert'
              >
                <div className='inline-block max-sm:mb-2'>
                  <svg
                    xmlns='http://www.w3.org/2000/svg'
                    className='w-5 fill-red-500 inline mr-4'
                    viewBox='0 0 32 32'
                  >
                    <path
                      d='M16 1a15 15 0 1 0 15 15A15 15 0 0 0 16 1zm6.36 20L21 22.36l-5-4.95-4.95 4.95L9.64 21l4.95-5-4.95-4.95 1.41-1.41L16 14.59l5-4.95 1.41 1.41-5 4.95z'
                      data-original='#ea2d3f'
                    />
                  </svg>
                  <strong className='font-bold text-sm font-medium'>
                    Error!
                  </strong>
                </div>
                <span className='block sm:inline text-sm mx-4 max-sm:ml-0 max-sm:mt-1'>
                  {errorAlert}
                </span>
                <svg
                  xmlns='http://www.w3.org/2000/svg'
                  className='w-7 hover:bg-red-200 rounded-md transition-all p-2 cursor-pointer fill-green-500 absolute right-4 top-1/2 -translate-y-1/2'
                  viewBox='0 0 320.591 320.591'
                >
                  <path
                    d='M30.391 318.583a30.37 30.37 0 0 1-21.56-7.288c-11.774-11.844-11.774-30.973 0-42.817L266.643 10.665c12.246-11.459 31.462-10.822 42.921 1.424 10.362 11.074 10.966 28.095 1.414 39.875L51.647 311.295a30.366 30.366 0 0 1-21.256 7.288z'
                    data-original='#000000'
                  />
                  <path
                    d='M287.9 318.583a30.37 30.37 0 0 1-21.257-8.806L8.83 51.963C-2.078 39.225-.595 20.055 12.143 9.146c11.369-9.736 28.136-9.736 39.504 0l259.331 257.813c12.243 11.462 12.876 30.679 1.414 42.922-.456.487-.927.958-1.414 1.414a30.368 30.368 0 0 1-23.078 7.288z'
                    data-original='#000000'
                  />
                </svg>
              </div>
            )}
          </div>

          <div className='font-[sans-serif] w-full mt-4 flex flex-col'>
            {finalDataTable && finalDataTable.length !== 0 && (
              <TableData data={finalDataTable} />
            )}
          </div>
        </div>
      </body>
    </>
  );
}

export default App;
