import * as React from 'react';
import { Web } from 'sp-pnp-js';
import '../assets/Admin.css';
import moment from 'moment';

import { MdDelete } from "react-icons/md";
import { FaEdit } from "react-icons/fa";
import { DefaultButton, PrimaryButton, TextField , Dropdown, IDropdownOption,Modal} from '@fluentui/react';
import { Panel ,PanelType  } from '@fluentui/react/lib/Panel';
import { Checkbox } from '@fluentui/react';
import Swal from 'sweetalert2';
import { RxCross2 } from "react-icons/rx";
import { FaAngleUp } from "react-icons/fa";
import { FaAngleDown } from "react-icons/fa";
import { useReactTable, flexRender, getCoreRowModel, getSortedRowModel, getFilteredRowModel, } from '@tanstack/react-table';

// interface TableProps {
//   tableData: any[];
//   fetchUserData: (month?: moment.Moment) => void; 
//   selectedMonth: moment.Moment;
//   chartSelectedsalesMonth?: string;
  
// }

type MedicalDetail = {
  BloodGroup: string;
  BP: number;
  HeartBeat: number;
  Sugar: number;
  Others: string;
};

type Paymenthistory={
  PaymentDate:string;
  Amount:string;
  Status:string;
  MembershipPlan:string
}
interface ClientItem {
  Id: number;
  FirstName: string;
  Title: string;
  FullName: string;
  Age: number ;
  Gender: string;
  Email: string;
  CellPhone: string;
  AadhaarNumber: number ;
  EndDate: string | null; 
  PaymentMode: string;
  AmountReceived: number;
  WorkAddress: string;
  JoiningDate: string | null ; 
  BillAmount: number;
  MedicalDetails: MedicalDetail[];
  PaymentHistory:Paymenthistory[];
  BMIDate: string | null; 
  NextBMIDueDate: string | null; 
  MembershipNo: string;
  MembershipPlan: string;
  DueDate: string|null;
  PaymentDue:string,
  Photo:string,
}
const Table = (tableData: any) =>  {
  let currentMonth = tableData?.selectedMonth
  const [data, setData] = React.useState<ClientItem[]>([]);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [isModalOpen, setIsModalOpen] = React.useState(false);
 const [editId, setEditId] = React.useState<number | null>(null);
 const [modalData, setModalData] = React.useState<{ PaymentDate: string; Amount: string; Status: string ,MembershipPlan:string}[]>([]);
 const [Createdname,setCreatedname] = React.useState();
  const [Createddate,setCreateddate] = React.useState();
  const [Modifiedname,setModifiedname] = React.useState();
  const [Modifiedate,setModifieddate] = React.useState();
  const [isActive, setIsActive] = React.useState(true);
  const [selectimage, setSelectedimage] = React.useState<File | null>(null);
  const [currimage, setCurrimage] = React.useState<File | null>(null);
  const [columnFilters, setColumnFilters] = React.useState<{ id: string; value: string }[]>([]);



const [newDetail, setNewDetail] = React.useState({
  PaymentDate: "",
  Amount: "",
  Status: "",
  MembershipPlan:""
});

const [inputValue, setInputValue] = React.useState({
    Id: 0,
    FirstName:'',
    Title:'',
    FullName: '',
    Age:0,
    Gender: 'Male',
    Email: '',
    AadhaarNumber:0,
    DueDate:'',
    CellPhone: '',
    WorkAddress: '',
    JoiningDate: '',
    EndDate:'',
    PaymentMode:'UPI',
    BillAmount: 0,
    AmountReceived:0,
    PaymentStatus:'',
    PaymentDue:'',
    PaymentHistory:[{
      PaymentDate:'',
      Amount:'',
      Status:"",
      MembershipPlan:""
      
    }],
    MedicalDetails:[{
    BloodGroup:"",
    BP:0,
    HeartBeat:0, 
    Sugar:0,
    Others:'',
}],
   BMIDate:'',
   NextBMIDueDate:'',
   MembershipNo: '',
   MembershipPlan: 'Monthly',
   Photo:""
   
  })
 
  const handleOpenModal = () => {
    if (inputValue?.PaymentHistory?.length) {
      setNewDetail(inputValue.PaymentHistory[0]); 
    } else {
      setNewDetail({ PaymentDate: "", Amount: "", Status: "", MembershipPlan:"" }); 
    }
    setIsModalOpen(true);
  };
  
  
const handleCloseModal = () => setIsModalOpen(false);

  
  const fetchAPIData = async () => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const res = await web.lists
        .getByTitle("Clients")
        .items.select(
          "Id",
          "FirstName",
          "Title",
          "FullName",
          "Age",
          "Gender",
          "Email",
          "CellPhone",
          "AadhaarNumber",
          "EndDate",
          "PaymentMode",
          "AmountReceived",
          "WorkAddress",
          "JoiningDate",
          "BillAmount",
          "MedicalDetails",
          "BMIDate",
          "NextBMIDueDate",
          "MembershipNo",
          "MembershipPlan",
          "PaymentHistory",
          "DueDate",
          "PaymentDue",
          "Author/ID",
          "Author/Title",
          "Editor/Id",
          "Editor/Title",
          "Modified",
          "Created",
          "isActive",
          "Photo"
        ).expand("Author,Editor")
        .top(4999)
        .get();
  
      res?.forEach((val:any) => {
        val.JoiningDate = val.JoiningDate
          ? moment(val.JoiningDate).format("DD/MM/YYYY")
          : null;
        val.EndDate = val.EndDate
          ? moment(val.EndDate).format("DD/MM/YYYY")
          : null;
        val.BMIDate = val.BMIDate
          ? moment(val.BMIDate).format("DD/MM/YYYY")
          : null;
        val.NextBMIDueDate = val.NextBMIDueDate
          ? moment(val.NextBMIDueDate).format("DD/MM/YYYY")
          : null;
  
       
        val.DueDate = calculateDueDate(val.JoiningDate, val.MembershipPlan);
      });
  
      setData(res);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };
  console.log(data)


  const calculateDueDate = (JoiningDate: any, MembershipPlan: any) => {
    if (!JoiningDate) return "";
  
    const date = moment(JoiningDate, "YYYY-MM-DD");

    if (!MembershipPlan) {
      return date.add(1, "month").format("YYYY-MM-DD");
    }
    switch (MembershipPlan.toLowerCase()) {
      case "Annual":
        return date.add(1, "year").format("YYYY-MM-DD");
      case "Bi-Annual":
        return date.add(6, "months").format("YYYY-MM-DD");
      case "Quarterly":
        return date.add(3, "months").format("YYYY-MM-DD");
      case "Monthly":
        return date.add(1, "month").format("YYYY-MM-DD");
      default:
        return "";
    }
  };
  



  const deleteFunction = async (id: number) => {
    try {
      const web = new Web('https://smalsusinfolabs.sharepoint.com/sites/F4S');
      await web.lists.getById('3A9C0B25-B14D-4277-99CB-D63FCFF5FD3F').items.getById(id).delete();
      
       tableData.fetchUserData(currentMonth); 
      // fetchUserData(selectedMonth); 
    } catch (error) {
      console.error('Error deleting item:', error);
    }
  };

  const uploadFileToLibrary = async (
    file: File,
    folderName: string,
    libraryName: string
  ) => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const fileName = file.name;
      const fileContent = await file.arrayBuffer();

      // Upload the file to libaray inside folder
      const uploadResult = await web
        .getFolderByServerRelativeUrl(
          `/sites/F4S/${libraryName}/${folderName}`
        )
        .files.add(fileName, fileContent, true);

      // Construct the file URL
      const fileUrl = `https://smalsusinfolabs.sharepoint.com${uploadResult.data.ServerRelativeUrl}`;
      return { fileUrl, fileName };
    } catch (error) {
      console.error("File upload failed:", error);
      throw error;
    }
  };

  
  const deleteimage = () => {
    setCurrimage(null);
    setSelectedimage(null);
  };
  

 
const handleDeleteTask = async (id: number) => {
  Swal.fire({
    title: "Are you sure?",
    text: "You won't be able to revert this!",
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#3085d6",
    cancelButtonColor: "#d33",
    confirmButtonText: "Yes, delete it!",

  }).then(async (result: any) => {
    if (result.isConfirmed) {
      try {
        setData(data.filter((task: any) => task?.Id !== id));

       await deleteFunction(id);

       
        Swal.fire({
          title: "Deleted!",
          text: "The item has been deleted successfully.",
          icon: "success",
        });
      } catch (error) {
        console.error("Data could not be deleted", error);

        
        Swal.fire({
          title: "Error!",
          text: "Something went wrong. The item could not be deleted.",
          icon: "error",
        });
      }
    }
  });
};
  
  React.useEffect(() => {
  fetchAPIData();
  }, []);

  // React.useEffect(() => {
  //   setData(tableData); 
  // }, [tableData]); 
  

 

  const updateDetails = async (id: number) => {
    let hyperlinkValue: any = null;

if (selectimage) {
  try {
    const { fileUrl, fileName } = await uploadFileToLibrary(
      selectimage,
      "UserImage",
      "UserDashboard"
    );

    if (fileUrl) {
      hyperlinkValue = { Url: fileUrl, Description: fileName };
    } else {
      console.error("File upload failed: No URL returned.");
    }
  } catch (error) {
    console.error("Error uploading file:", error);
  }
}
    
    try {
     const web = new Web('https://smalsusinfolabs.sharepoint.com/sites/F4S');
      await web.lists.getById('3A9C0B25-B14D-4277-99CB-D63FCFF5FD3F').items.getById(id).update({
         FirstName: inputValue.FirstName,
          Title: inputValue.Title,
          FullName: inputValue.FullName,
          Email: inputValue.Email,
          Age: inputValue?.Age?inputValue.Age:0,
          BillAmount:inputValue?.BillAmount?inputValue.BillAmount:0,
          Gender: inputValue.Gender,
          AadhaarNumber: inputValue?.AadhaarNumber ? inputValue?.AadhaarNumber:0,
          CellPhone: inputValue.CellPhone,
          PaymentDue:inputValue.PaymentDue,
          MembershipPlan:inputValue.MembershipPlan,
          WorkAddress: inputValue.WorkAddress,
          JoiningDate: inputValue.JoiningDate
          ? new Date(inputValue.JoiningDate).toISOString()
          : null,
          EndDate: inputValue.EndDate
          ? new Date(inputValue.EndDate).toISOString()
          : null,
          PaymentMode:inputValue.PaymentMode,
          AmountReceived: inputValue.AmountReceived ? inputValue.AmountReceived:null,
          MedicalDetails: JSON.stringify(inputValue?.MedicalDetails),
          PaymentHistory:JSON.stringify(inputValue.PaymentHistory),
          BMIDate: inputValue.BMIDate
          ? new Date(inputValue.BMIDate).toISOString()
          : null,
        NextBMIDueDate: inputValue.NextBMIDueDate
          ? new Date(inputValue.NextBMIDueDate).toISOString()
          : null,
          DueDate: inputValue.DueDate
          ? new Date(inputValue.DueDate).toISOString()
          : null,
          isActive: isActive,
          Photo: hyperlinkValue || null,
        });
     
      setIsPanelOpen(false);
      setEditId(null)
     
    
      tableData.fetchUserData(currentMonth); 

      // fetchUserData(selectedMonth); 
    } catch (error) {
      console.error('Error updating item:', error);

    }
  };
 const handleEditTask = (task: any) => {
  setEditId(task.Id);
   const parsedRowData = {
        ...task,
        JoiningDate: task.JoiningDate
      ? moment(task.JoiningDate, "DD/MM/YYYY").format("YYYY-MM-DD")
      : null,
    EndDate: task.EndDate
      ? moment(task.EndDate, "DD/MM/YYYY").format("YYYY-MM-DD")
      : null,
    BMIDate: task.BMIDate
      ? moment(task.BMIDate, "DD/MM/YYYY").format("YYYY-MM-DD")
      : null,
    NextBMIDueDate: task.NextBMIDueDate
      ? moment(task.NextBMIDueDate, "DD/MM/YYYY").format("YYYY-MM-DD")
      : null,
     DueDate: task.DueDate
      ? moment(task.DueDate, "DD/MM/YYYY").format("YYYY-MM-DD")
      : null,
      
        PaymentHistory: task.PaymentHistory && typeof task.PaymentHistory === "string"
        ? JSON.parse(task.PaymentHistory)
        : [
            {
              PaymentDate: task.PaymentDate
              ? moment(task.PaymentDate, "DD/MM/YYYY").format("YYYY-MM-DD")
              : null,
                Amount: "",
                Status: "",
                MembershipPlan:""
            },
        ],
        isActive: task.isActive,
        Photo: task.Photo?.Url || "",
       
    
    };
    setCurrimage(task?.Photo?.Url ? task?.Photo?.Url : null);
    setCreatedname(task.Author?.Title || "");
    setCreateddate(task.Created || "");
    setModifiedname(task.Editor?.Title || "");
    setModifieddate(task.Modified || "");

    console.log("Parsed data:", parsedRowData);
    setInputValue(parsedRowData); 
    setIsPanelOpen(true); 
    setIsActive(task.isActive);
    // setEditId(null);
};

const handleAddRow = () => {
  if (newDetail.PaymentDate && newDetail.Amount && newDetail.Status) {
    setInputValue((prev) => ({
      ...prev,
      PaymentHistory: [...(prev.PaymentHistory || []), { ...newDetail }], // Ensure PaymentHistory is always an array
    }));

    // Reset the newDetail state after adding the row
    setNewDetail({
      PaymentDate: "",
      Amount: "",
      Status: "",
      MembershipPlan: "",
    });
  } else {
    console.warn("Fill all payment details before adding a row.");
  }
};

const handleSaveTask = async () => {
  try {
    if (editId !== null) {
     await updateDetails(editId);
      Swal.fire({
       
        text: "You have successfully updated items!",
        icon: "success",
      });
    } else {
      await handleAddTask();
      Swal.fire({
       
        text: "You have successfully added items!",
        icon: "success",
      });
    }

    
    setInputValue((prev) => ({
      ...prev,
      PaymentHistory: [...prev.PaymentHistory, ...modalData],
    }));

    // Close the panel
    setIsPanelOpen(false);
  } catch (error) {
    console.error("Error saving task:", error);
    Swal.fire({
      title: "Error!",
      text: "An error occurred while saving the task.",
      icon: "error",
    });
  }
  setIsPanelOpen(false);
};




const handleSaveModal = () => {
 
  setInputValue((prev) => ({
    ...prev,
    PaymentHistory: [...prev.PaymentHistory, ...modalData],
  }));

 
  setModalData([]);

 
  setIsModalOpen(false);
};



  const handleClosePanel = () => {
    setInputValue({
      Id: 0,
      FirstName: "",
      Title: "",
      FullName: "",
      Age: 0,
      DueDate:'',
      Gender: "",
      Email: "",
      AadhaarNumber: 0,
      CellPhone: "",
      WorkAddress: "",
      JoiningDate: "",
      EndDate: "",
      PaymentMode: "",
      BillAmount: 0,
      AmountReceived: 0,
      PaymentStatus:"",
      PaymentHistory:[{
        PaymentDate:'',
        Amount:"",
        Status:"",
       MembershipPlan:"",
        
      }],
      MedicalDetails: [
        {
          BloodGroup: "",
          BP: 0,
          HeartBeat: 0,
          Sugar: 0,
          Others: "",
        },
      ],
      BMIDate: "",
      NextBMIDueDate: "",
      MembershipNo: "",
      MembershipPlan: "",
      PaymentDue:"",
      Photo:""
    });
  
   
    setIsPanelOpen(false);
  };

  const handleCheckboxChange = (ev:any, checked:any) => {
    setIsActive(checked); // Update the isActive state when checkbox is clicked
  };

  
  const handleInputChange = (
    field: keyof ClientItem,
    value: string | number | FileList | null,
    e?: React.ChangeEvent<HTMLInputElement>
  ) => {
    setInputValue((prev) => {
      const updatedState = { ...prev, [field]: value };
  
      // Auto-update FullName based on FirstName and Title
      if (field === "FirstName" || field === "Title") {
        updatedState.FullName = `${updatedState.FirstName || ""} ${updatedState.Title || ""}`.trim();
      }
  
      // Auto-calculate DueDate based on JoiningDate and MembershipPlan
      if (field === "JoiningDate" || field === "MembershipPlan") {
        updatedState.DueDate = calculateDueDate(
          updatedState.JoiningDate as string,
          updatedState.MembershipPlan as string
        );
      }
  
     
      if (field === "Photo" && e?.target.files && e.target.files[0]) {
        setSelectedimage(e.target.files[0]);
        updatedState.Photo = URL.createObjectURL(e.target.files[0]); // Show image preview
      }
  
    
      if (field === "BillAmount" || field === "AmountReceived" || field === "PaymentDue") {
        const paymentDue = Number(updatedState.PaymentDue || 0);
        const amountReceived = Number(updatedState.AmountReceived || 0);
        updatedState.BillAmount = Math.round(paymentDue + amountReceived);
      }
  
      return updatedState;
    });
  };
  


  const handleMedicalDetailsChange = (field: keyof MedicalDetail, value: any) => {
    setInputValue((prev) => ({
      ...prev,
      MedicalDetails: [
        {
          ...prev.MedicalDetails[0],
          [field]: value,
        },
      ],
    }));
  };

  const handlepaymentdetailChange = (field: keyof Paymenthistory, value: any) => {
    setNewDetail((prev) => ({
      ...prev,
      [field]: value,
    }));
  };
  
const genderOptions: IDropdownOption[] = [
    { key: 'Male', text: 'Male' },
    { key: 'Female', text: 'Female' },
    { key: 'Others', text: 'Others' },
  ];
  const MembershipPlansOptions: IDropdownOption[] = [
    { key: 'Annual', text: 'Annual' },
    { key: 'Bi-Annual', text: 'Bi-Annual' },
    { key: 'Monthly', text: 'Monthly' },
    { key: 'Quarterly', text: 'Quarterly' },
  ];

  const paymentModeOptions: IDropdownOption[] = [
    { key: 'UPI', text: 'UPI' },
    { key: 'Credit Card', text: 'Credit Card' },
    { key: 'Net Banking', text: 'Net Banking' },
    { key: 'Cash', text: 'Cash' },
  ];

  const statusOptions:IDropdownOption[] = [
    { key: "Paid", text: "Paid" },
    { key: "Unpaid", text: "Unpaid" },
    { key: "Pending", text: "Pending" },
  ];

  const getNextMembershipNo = async () => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const existingData = await web.lists
        .getByTitle("Clients")
        .items.select("MembershipNo")
        .top(4999)
        .get();
  
     const membershipNos = existingData
        .map((item: { MembershipNo: string; }) => parseInt(item.MembershipNo))
        .filter((num: number) => !isNaN(num)); 
  
      const uniqueMembershipNos = membershipNos.filter(
        (num: any, index: any, arr: string | any[]) => arr.indexOf(num) === index
      );
   const maxMembershipNo = Math.max(...uniqueMembershipNos);
   return maxMembershipNo === -Infinity ? 1 : maxMembershipNo + 1;
    } catch (error) {
      console.error("Error calculating next MembershipNo:", error);
      throw error; 
    }
  };
  
  const handleAddTask = async () => {
    setEditId(null);
    let hyperlinkValue: any = null;
    if (selectimage) {
      const { fileUrl, fileName } = await uploadFileToLibrary(
        selectimage,
        "UserDashboard",
        "UserImage"
      );
      hyperlinkValue = { Url: fileUrl, Description: fileName };
    }
    try {

      const nextMembershipNo = await getNextMembershipNo();
      const postData = {
          FirstName: inputValue.FirstName,
          Title: inputValue.Title,
          FullName: inputValue.FullName,
          Email: inputValue.Email,
          Age: inputValue?.Age?inputValue.Age:0,
          PaymentDue:inputValue.PaymentDue,
          BillAmount: inputValue?.BillAmount?inputValue.BillAmount:0,
          Gender: inputValue.Gender,
          AadhaarNumber: inputValue?.AadhaarNumber ? inputValue?.AadhaarNumber:0,
          CellPhone: inputValue.CellPhone,
          WorkAddress: inputValue.WorkAddress,
          JoiningDate: inputValue.JoiningDate
          ? new Date(inputValue.JoiningDate).toISOString()
          : null,
        EndDate: inputValue.EndDate
          ? new Date(inputValue.EndDate).toISOString()
          : null,
          MembershipNo: nextMembershipNo.toString(),
          MembershipPlan:inputValue.MembershipPlan,
        PaymentMode:inputValue.PaymentMode,
          AmountReceived: inputValue.AmountReceived ? inputValue.AmountReceived:null,
          MedicalDetails: JSON.stringify(inputValue?.MedicalDetails),
          PaymentHistory:JSON.stringify(inputValue.PaymentHistory),
          BMIDate: inputValue.BMIDate
          ? new Date(inputValue.BMIDate).toISOString()
          : null,
        NextBMIDueDate: inputValue.NextBMIDueDate
          ? new Date(inputValue.NextBMIDueDate).toISOString()
          : null,
          DueDate: inputValue.DueDate
          ? new Date(inputValue.DueDate).toISOString()
          : null,
          Photo: hyperlinkValue || null,
        };
        const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
        let res = await web.lists
          .getById("3A9C0B25-B14D-4277-99CB-D63FCFF5FD3F")
          .items.add(postData);
  
        console.log("Added item response:", res);
        //tableData.push(res)
        setInputValue({
          Id: 0,
          FirstName:"",
          Title:"",
          FullName: "",
          Age: 0,
          Gender: "",
          Email: "",
          DueDate:'',
          AadhaarNumber: 0,
          CellPhone: "",
          WorkAddress: "",
          JoiningDate: "",
          EndDate: "",
          PaymentMode:"",
          BillAmount: 0,
          AmountReceived: 0,
          PaymentStatus:"",
          MedicalDetails:[{
            BloodGroup:"",
            BP:0,
            HeartBeat:0, 
            Sugar:0,
            Others:'',
        
         }],
         PaymentHistory:[{
          PaymentDate:'',
          Amount:"",
          Status:"",
          MembershipPlan:""
          
        }],
         BMIDate:"",
         NextBMIDueDate:"",
         MembershipNo:"",
         MembershipPlan:"",
         PaymentDue:"",
         Photo:""
        });
     
       tableData.fetchUserData(currentMonth); 
      // fetchUserData(selectedMonth); 
      } catch (error) {
        console.error("Error adding item:", error);
      }
    
  };

 const handleNewTask = () => {
      
        setInputValue({
          Id: 0,
          FirstName: "",
          Title: "",
          FullName: "",
          Age: 0,
          Gender: "",
          Email: "",
          DueDate: "",
          AadhaarNumber: 0,
          CellPhone: "",
          WorkAddress: "",
          PaymentDue:"",
          JoiningDate: "",
          EndDate: "",
          PaymentMode: "",
          BillAmount: 0,
          AmountReceived: 0,
          PaymentStatus: "",
          MedicalDetails: [{
            BloodGroup: "",
            BP: 0,
            HeartBeat: 0,
            Sugar: 0,
            Others: "",
          }],
          PaymentHistory: [{
            PaymentDate: '',
            Amount: "",
            Status: "",
            MembershipPlan:""
          }],
          BMIDate: "",
          NextBMIDueDate: "",
          MembershipNo: "",
          MembershipPlan: "",
          Photo:""
        });
        setEditId(null); 
        setIsPanelOpen(true); 
      };
     
      const onRenderFooterContent = React.useCallback(
        () => (
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
            {editId != null && (
              <div>
                <div>
                  Created{" "}
                  <span style={{ color: "blue" }}>
                    {Createddate ? moment(Createddate).format("DD MMM YYYY") : "N/A"}
                  </span>{" "}
                  by{" "}
                  <span style={{ color: "blue" }}>
                    {Createdname || "N/A"}
                  </span>
                </div>
                <div>
                  Last modified{" "}
                  <span style={{ color: "blue" }}>
                    {Modifiedate ? moment(Modifiedate).format("DD MMM YYYY") : "N/A"}
                  </span>{" "}
                  by{" "}
                  <span style={{ color: "blue" }}>
                    {Modifiedname || "N/A"}
                  </span>
                </div>
              </div>
            )}
      
            <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
              {editId != null && (
                <a
                href="#"
                style={{ textDecoration: "none", }}
                onClick={(e) => {
                  e.preventDefault(); 
                  window.open(
                    `https://smalsusinfolabs.sharepoint.com/sites/F4S/Lists/Clients/EditForm.aspx?ID=${editId}`,
                    "_blank"
                  );
                }}
              >
                Open out-of-the-box form
              </a>
              
              )}
              <PrimaryButton onClick={handleSaveTask}>Save</PrimaryButton>
              <DefaultButton onClick={() => handleClosePanel()}>Cancel</DefaultButton>
            </div>
          </div>
        ),
        [handleSaveTask, handleClosePanel, Createddate, Createdname, Modifiedate, Modifiedname, editId]
      );
      
      const isDiscontinued = (row: any) => {
        if (!row?.EndDate) return false;
    
        const discontinueDate =  moment(row.EndDate, "DD/MM/YYYY");
        const isOver = moment().isAfter(discontinueDate, "day");
        const isInactive = row.isActive === false;
    
        return isOver && isInactive;
      };


      const columns = [
        {
          header: "M.No",
          accessorKey: "MembershipNo",
        },
        {
          header: "Name",
          accessorKey: "FullName",
          cell: ({ row }: any) => {
            const item = row.original;
            return (
              <a
                href="#"
                onClick={(e) => {
                  e.preventDefault();
                  window.open(
                    item.Id
                      ? `https://smalsusinfolabs.sharepoint.com/sites/F4S/SitePages/MyDashboard.aspx?userId=${item.Id}`
                      : `https://smalsusinfolabs.sharepoint.com/sites/F4S/SitePages/MyDashboard.aspx`,
                    "_blank"
                  );
                }}
                style={{ color: "#007bff", textDecoration: "none" }}
              >
                {item.FullName}
              </a>
            );
          },
        },
        {
          header: "Gender",
          accessorKey: "Gender",
        },
        {
          header: "Mobile No.",
          accessorKey: "CellPhone",
        },
        {
          header: "Joining Date",
          accessorKey: "JoiningDate",
        },
        {
          header: "End Date",
          accessorKey: "EndDate",
          cell: ({ row }: any) => (isDiscontinued(row.original) ? row.original.EndDate : ""),
        },
        {
          header: "Due Date",
          accessorKey: "DueDate",
          cell: ({ row }: any) => (!isDiscontinued(row.original) ? row.original.DueDate : ""),
        },
        {
          header: "Amount",
          accessorKey: "BillAmount",
        },
        {
          header: "Received",
          accessorKey: "AmountReceived",
        },
        {
          header: "Pending",
          accessorKey: "PaymentDue",
        },
        {
          header: "Plan",
          accessorKey: "MembershipPlan",
        },
        {
          header: " ",
          cell: ({ row }: any) => {
            const item = row.original;
            return (
              <div style={{ display: "flex", alignItems: "center",  }}>
                <span
                  className="btn"
                  onClick={() => handleEditTask(item)}
                >
                  <FaEdit style={{ color: "green" }} />
                </span>
                <span
                  className="btn"
                  onClick={() => handleDeleteTask(item.Id)}
                >
                  <MdDelete style={{ color: "red" }} />
                </span>
              </div>
            );
          },
        },
      ];
      
      const tabledata= useReactTable({
        data: tableData?.tableData || [],
        columns,
        state: {
          columnFilters,
          sorting: [{ id: "MembershipNo", desc: true }], 
        },
        onColumnFiltersChange: setColumnFilters,
        getCoreRowModel: getCoreRowModel(),
        getSortedRowModel: getSortedRowModel(),
        getFilteredRowModel: getFilteredRowModel(),
        initialState: {
          sorting: [{ id: "MembershipNo", desc: true }], 
        },
      });
      
        

     

  return (
    <>
    <div className='text-end w-100'>
      <button className="border-0 my-2 px-3 py-2 rounded-2 btn btn-primary" 
        onClick={handleNewTask}
      >
        New Registration
      </button>
      </div>
      <Panel
      isOpen={isPanelOpen}
      onDismiss={handleClosePanel}
      headerText="Form Details"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={onRenderFooterContent}
      isFooterAtBottom={true}
      type={PanelType.custom} 
      styles={{
       main: {
      width: "80vw", 
      maxWidth: "1200px", 
      },
        content: {
        padding: "24px", 
       },
     }}
   >
        <div className="container">
          {/* Row 1 */}
          <div className="row align-items-center mb-4">
          <div className="col">
              <h6> MembershipNo</h6>
              <TextField
               value={inputValue.MembershipNo}
                disabled
              />
            </div>
         <div className="col">
          <h6>First Name</h6>
            <TextField
      placeholder="Enter First Name"
      value={inputValue.FirstName}
      onChange={(e) => handleInputChange("FirstName", e.currentTarget.value)}
      required
      errorMessage={!inputValue?.FirstName ? "First name  is required" : ""}
      autoComplete='off'
    />
  </div>
  
  <div className="col">
    <h6>Last Name</h6>
    <TextField
      placeholder="Enter Last Name"
      value={inputValue.Title}
      onChange={(e) => handleInputChange("Title", e.currentTarget.value)}
       autoComplete='off'
    />
  </div>
  
  <div className="col">
    <h6>Full Name</h6>
    <TextField
      placeholder="Enter Full Name"
      value={inputValue.FullName}

      onChange={(e) => handleInputChange("FullName", e.currentTarget.value)}
       autoComplete='off'
    />
  </div>
            </div>


          {/* Row 2 */}
          <div className="row align-items-center mb-4">
          <div className="col">
           <h6>Age</h6>
          <TextField
            placeholder="Enter Age"
           type="number"
          value={inputValue.Age != null ? inputValue.Age?.toString() : ""} 
           onChange={(e) => {
          const parsedValue = parseInt(e.currentTarget.value, 10);
          handleInputChange("Age", isNaN(parsedValue) ? "" : parsedValue);
         }}
          autoComplete='off'
           />
      </div>

          <div className="col">
              <h6>Gender</h6>
              <Dropdown
                selectedKey={inputValue.Gender}
                options={genderOptions}
                onChange={(e, option) => handleInputChange("Gender", option?.key || "")}
                 
              />
            </div>
            <div className="col">
              <h6>Email</h6>
              <TextField
                placeholder="Enter Email"
                value={inputValue.Email}
                onChange={(e) => handleInputChange("Email", e.currentTarget.value)}
                 autoComplete='off'
              />
            </div>
            <div className="col">
              <h6>Aadhaar Number</h6>
              <TextField
                placeholder="Enter Aadhaar Number"
                type="number"
                value={inputValue.AadhaarNumber !== undefined ? inputValue.AadhaarNumber?.toString() : ""}
                onChange={(e) => handleInputChange("AadhaarNumber", Number(e.currentTarget.value))}
                 autoComplete='off'
              />
            </div>
            
          </div>

          {/* Row 3 */}
          <div className="row align-items-center mb-4">
        
           
          <div className="col">
              <h6>Mobile No</h6>
              <TextField
                placeholder="Enter Cell Phone"
                value={inputValue.CellPhone}
                onChange={(e) => handleInputChange("CellPhone", e.currentTarget.value)}
                 autoComplete='off'
              />
            </div>
            <div className="col">
           <h6> Address</h6>
           <textarea
           placeholder="Enter Work Address"
           value={inputValue.WorkAddress}
          onChange={(e) => handleInputChange("WorkAddress", e.target.value)}
           rows={2} 
           style={{
           width: '100%', 
           padding: '8px',
            fontSize: '16px',
            }}
        ></textarea>
               </div>
               <div className="col">
              <h6>Joining Date</h6>
              <TextField
                placeholder="Enter Joining Date"
                type="date"
                value={inputValue.JoiningDate}
                onChange={(e) => handleInputChange("JoiningDate", e.currentTarget.value)}
                required
                errorMessage={!inputValue?.JoiningDate ? "Joining date is required" : ""}
              />
            </div>
            <div className="col">
              <h6>End Date</h6>
              <TextField
                placeholder="Enter End Date"
                type="date"
                value={inputValue.EndDate}
                onChange={(e) => handleInputChange("EndDate", e.currentTarget.value)}
              />
            </div>

           
           
          </div>

          {/* Row 4 */}
          <div className="row align-items-center mb-4">
          
         
            <div className="col">
              <h6>Payment Mode</h6>
              <Dropdown
                selectedKey={inputValue.PaymentMode}
                options={paymentModeOptions}
                onChange={(e, option) => handleInputChange("PaymentMode", option?.key || "")}
              />
            </div>

            <div className="col">
       <h6>Bill Amount</h6>
        <TextField
        placeholder="Enter Bill Amount"
         type="number"
        value={inputValue.BillAmount !== undefined ? inputValue.BillAmount?.toString() : ""}
        onChange={(e) => {
        const value = Number(e.currentTarget.value);
        handleInputChange("BillAmount", value);
        
      }}
        />
          </div>

          <div className="col">
         <h6>Amount Received</h6>
           <TextField
            placeholder="Enter Amount Received"
            type="number"
            value={inputValue.AmountReceived !== undefined ? inputValue.AmountReceived?.toString() : ""}
           onChange={(e) => {
        const value = Number(e.currentTarget.value);
        handleInputChange("AmountReceived", value);
        // updatePaymentStatus(inputValue.BillAmount, value);
         }}
        />
            </div>
            <div className="col">
              <h6>BMI date</h6>
              <TextField
                placeholder="Enter BMIDate"
                type='date'
                value={inputValue.BMIDate}
                onChange={(e) => handleInputChange("BMIDate", e.currentTarget.value)}
              />
            </div>
           
          </div>
          {/* Row 5 */}
          <div className="row align-items-center mb-4">
          <div className="col-12">
          <h6>Medical Details</h6>
          
         <div
         style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(5, 1fr)', 
         gap: '16px', 
        width: '100%',
       }}
     >
  <Dropdown
    placeholder="Select Blood Group"
    selectedKey={
      inputValue.MedicalDetails && inputValue.MedicalDetails.length > 0 
        ? inputValue.MedicalDetails[0].BloodGroup 
        : null
    }
    options={[
      { key: 'A+', text: 'A+' },
      { key: 'A-', text: 'A-' },
      { key: 'B+', text: 'B+' },
      { key: 'B-', text: 'B-' },
      { key: 'O+', text: 'O+' },
      { key: 'O-', text: 'O-' },
      { key: 'AB+', text: 'AB+' },
      { key: 'AB-', text: 'AB-' },
    ]}
    onChange={(e, option) =>
      handleMedicalDetailsChange('BloodGroup', option?.key || '')
    }
    styles={{ dropdown: { width: '100%' } }}
  />

  <TextField
    placeholder="Enter BP"
    type="number"
    value={
      inputValue.MedicalDetails && Array.isArray(inputValue.MedicalDetails) && inputValue.MedicalDetails[0]
        ? inputValue.MedicalDetails[0]?.BP?.toString() || '' 
        : ''
    }
    onChange={(e) =>
      handleMedicalDetailsChange('BP', Number(e.currentTarget.value) || 0)
    }
    styles={{ root: { width: '100%' } }}
  />

  <TextField
    placeholder="Enter Heart Beat"
    type="number"
    value={
      inputValue.MedicalDetails?.[0]?.HeartBeat
        ? inputValue.MedicalDetails[0].HeartBeat.toString()
        : ''
    }
    onChange={(e) =>
      handleMedicalDetailsChange('HeartBeat', Number(e.currentTarget.value) || 0)
    }
    styles={{ root: { width: '100%' } }}
  />

  <TextField
    placeholder="Enter Sugar"
    type="number"
    value={inputValue.MedicalDetails?.[0]?.Sugar ? inputValue.MedicalDetails[0].Sugar?.toString() : ''}
    onChange={(e) =>
      handleMedicalDetailsChange('Sugar', Number(e.currentTarget.value) || 0)
    }
    styles={{ root: { width: '100%' } }}
  />

  <TextField
    placeholder="Others"
    value={inputValue.MedicalDetails?.[0]?.Others || ''}
    onChange={(e) =>
      handleMedicalDetailsChange('Others', e.currentTarget.value)
    }
    styles={{ root: { width: '100%' } }}
  />
           </div>

      </div>
      </div>
     
      
             {/* row-6 */}
             <div className="row align-items-center mb-4">
             <div className="col">
              <h6>NextBMIDueDate</h6>
              <TextField
                placeholder="Enter NextBMIDueDate"
                 type='date'
                 value={inputValue.NextBMIDueDate}
                onChange={(e) => handleInputChange("NextBMIDueDate", e.currentTarget.value)}
              />
            </div>
             
             <div className="col">
              <h6>MembershipPlan</h6>
              <Dropdown
                selectedKey={inputValue.MembershipPlan || null}
                options={MembershipPlansOptions}
                onChange={(e, option) => handleInputChange("MembershipPlan", option?.key || "")}
              />
            </div>
            <div className="col">
            <h6>Payment Due</h6>
            <TextField
            placeholder="PaymentDue"
            type="text"
            value={inputValue.PaymentDue}
            onChange={(e) => handleInputChange("PaymentDue", e.currentTarget.value)}
              />
            </div>
            <div className="col">
              <h6>DueDate</h6>
              <TextField
                placeholder="DueDate"
                 type='date'
                 value={inputValue.DueDate}
                onChange={(e) => handleInputChange("DueDate", e.currentTarget.value)}
              />
            </div>

              </div>
           <div className='row align-items-center mb-4'>
           <div className="col">
            <h6>Payment Status</h6>
            <div style={{ display: "flex", alignItems: "center" }}>
            <TextField
              readOnly
             value={inputValue.PaymentStatus || ""}
            placeholder="Payment Status"
            styles={{
           root: {
           flexGrow: 1, 
           },
            }}
           />
          <FaEdit
           style={{
           marginLeft: "8px",
           cursor: "pointer",
           color: "gray",
           }}
           onClick={handleOpenModal}
           />
         </div>

          </div>
          
          {editId !== null && (
         <div className="col-lg-4">
         <h6>Active</h6>
         <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
           <Checkbox label="Yes" checked={isActive} onChange={handleCheckboxChange} />
           <Checkbox label="No" checked={!isActive} onChange={handleCheckboxChange} />
         </div>
       </div>
       
       )}
       

<div className="col-lg-4">
                <div className="form-group m-2">
                  <h6>Upload Photo</h6>
                  {currimage && editId ? (
                    <span>
                      <img
                        src={inputValue?.Photo}
                        alt="Uploaded Preview"
                        style={{
                          width: "50px",
                          height: "50px",
                          objectFit: "cover",
                          border: "1px solid black",
                          borderRadius: "50%",
                        }}
                      />
                      <RxCross2
                        onClick={deleteimage}
                        style={{
                          fontSize: "16px",
                          marginTop: "-46px",
                          marginLeft: "-5px",
                          color: "black",
                        }}
                      />
                    </span>
                  ) : (
                    <input
                    type="file"
                    className="form-control"
                    name="Photo"
                    onChange={(e) => handleInputChange("Photo", e.target.files, e)}
                    autoComplete="off"
                  />
                  
                  )}
                </div>
              </div>
         


         
         

       </div>
       
        

        

        </div>


              <Modal
           isOpen={isModalOpen}
           onDismiss={handleCloseModal}
          isBlocking={false}
          styles={{ main: { width: "600px" } }}
         >
            

       <div style={{ padding: "20px", background: "white", borderRadius: "4px" }}>
       <h5 style={{ marginBottom: "20px" }}>Add Payment Details</h5>
       <div style={{ marginTop: "20px" }}>
         <TextField
        label="Payment Date"
        type="date"
        value={newDetail.PaymentDate || ""}
        placeholder="Enter Date"
        onChange={(e) =>
          handlepaymentdetailChange("PaymentDate", e.currentTarget.value)
        }
      />
      <TextField
        label="Amount"
        type="number"
         value={newDetail.Amount || ""}
          placeholder="Enter Amount"
        onChange={(e) =>
          handlepaymentdetailChange("Amount", e.currentTarget.value)
        }
      />
       <Dropdown
     label="Status"
     placeholder="Select a status"
  options={statusOptions}
  selectedKey={newDetail.Status || ""}
  onChange={(e, option) => handlepaymentdetailChange("Status",option?.key || '')}
/>;
<Dropdown
     label="MembershipPlan"
     placeholder="Select a Plan"
  options={MembershipPlansOptions}
  selectedKey={newDetail.MembershipPlan || ""}
  onChange={(e, option) => handlepaymentdetailChange("MembershipPlan",option?.key || '')}
/>
      <PrimaryButton
        onClick={handleAddRow}
        styles={{ root: { marginTop: 16 } }}
        disabled={
          !newDetail.PaymentDate || !newDetail.Amount || !newDetail.Status
        }
      >
        Add Row
      </PrimaryButton>
    </div>


    <table className="table">
      <thead>
        <tr>
          <th>Payment Date</th>
          <th>Amount</th>
          <th>Status</th>
          <th>Membership Plan</th>
        </tr>
      </thead>
      <tbody>
        {inputValue.PaymentHistory?.map((detail, index) => (
          <tr key={index}>
            <td>
          {detail.PaymentDate ? moment(detail.PaymentDate).isValid() ? moment(detail.PaymentDate).format("DD/MM/YYYY") : "" : ""}
        </td>
            <td>{detail.Amount}</td>
            <td>{detail.Status}</td>
            <td>{detail.MembershipPlan}</td>
          </tr>
        ))}
      </tbody>
    </table>
   

    <div style={{ marginTop: "20px", display: "flex", justifyContent: "flex-end" }}>
      <PrimaryButton
        onClick={handleSaveModal}
        styles={{ root: { marginRight: 8 } }}
      >
        Save
      </PrimaryButton>
      <DefaultButton onClick={handleCloseModal}>Cancel</DefaultButton>
    </div>
       </div>
              </Modal>
          </Panel>
          <div className="m-3 mb-3 bg-light" style={{ maxHeight: "500px", overflowY: "auto" }}>
  <table
    className="table-striped table-bordered table-hover bg-light"
    style={{
      borderCollapse: "collapse",
      width: "100%",
      border: "1px solid #ddd",
    }}
  >
    <thead>
      {tabledata.getHeaderGroups().map((headerGroup) => (
        <tr key={headerGroup.id}>
          {headerGroup.headers
            .filter((header) => header.column.id !== "Status") 
            .map((header) => (
              <th key={header.id} style={{ padding: "8px", border: "1px solid #ddd" }}>
                <div className="position-relative">
                
                  {header.column.getCanFilter() ? (
                    <div className="position-relative">
                      <input
                        type="text"
                        placeholder={` ${header.column.columnDef.header}`}
                        value={(header.column.getFilterValue() as string) || ""}
                        onChange={(e) => header.column.setFilterValue(e.target.value)}
                        className="form-control form-control-sm mt-1"
                        style={{ 
                          fontSize: "12px", 
                          paddingRight: "24px", 
                          height: "30px" 
                        }}
                      />
                    
                      <span
                        className="position-absolute"
                        style={{
                          right: "6px",
                          top: "50%",
                          transform: "translateY(-50%)",
                          color: "#97999b",
                          cursor: "pointer",
                          display: "flex",
                          flexDirection: "column",
                          alignItems: "center",
                          fontSize: "10px", 
                          gap: "2px",
                        }}
                        onClick={header.column.getToggleSortingHandler()}
                      >
                        <FaAngleUp style={{ fontWeight: "normal" }} />
                        <FaAngleDown style={{ fontWeight: "normal" }} />
                      </span>
                    </div>
                  ) : (
                    flexRender(header.column.columnDef.header, header.getContext())
                  )}
                </div>
              </th>
            ))}
        </tr>
      ))}
    </thead>
    <tbody>
  {tabledata.getRowModel().rows.map((row) => (
   <tr 
   key={row.id} 
   className={row.original.highlightColor === "grey" ? "grey-row" : ""}
 >
 
 
 
  
 
      {row.getVisibleCells()
        .filter((cell) => cell.column.id !== "Status") // Remove "Status" column
        .map((cell) => (
          <td key={cell.id} style={{ padding: "8px", border: "1px solid #ddd" }}>
            {flexRender(cell.column.columnDef.cell, cell.getContext())}
          </td>
        ))}
    </tr>
  ))}
</tbody>

  </table>
</div>


      
    </>
  );
}

export default Table;

