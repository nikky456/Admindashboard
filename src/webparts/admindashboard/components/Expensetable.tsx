import * as React from 'react';
import { Web } from 'sp-pnp-js';
import { DefaultButton, PrimaryButton, TextField , Dropdown, IDropdownOption} from '@fluentui/react';
import { Panel, PanelType  } from '@fluentui/react/lib/Panel';
import moment from 'moment';
import Swal from 'sweetalert2';
import "bootstrap/dist/css/bootstrap.min.css";
import { MdDelete } from "react-icons/md";
 import { FaEdit } from "react-icons/fa";

 interface TableProps{
  expenseData :any[];
  fetchExpenseAPIData:(month?: moment.Moment) => void; 
  selectedMonth: moment.Moment;
  chartselectedMonth: string | null;
 }

type Expense = {
  Id: number;
  ExpenseDate: string; 
  PaymentMode: string;
  ExpenseType: string;
  Amount: number;
  AmountType:string;
  Comment:string;
};


const Expensetable:React.FC<TableProps>= ( {expenseData,fetchExpenseAPIData,selectedMonth, chartselectedMonth} ) => {
  // let currentMonth =expenseData?.selectedMonth
  const [data, setData] = React.useState<Expense[]>([]);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [editId, setEditId] = React.useState(null);
  const [Createdname,setCreatedname] = React.useState();
  const [Createddate,setCreateddate] = React.useState();
  const [Modifiedname,setModifiedname] = React.useState();
  const [Modifiedate,setModifieddate] = React.useState();

  const [inputValue, setInputValue] = React.useState({
    ExpenseDate:'',
    PaymentMode:'UPI',
    ExpenseType:'',
    Amount:0,
    AmountType:'Expense',
    Comment:""
   
  })

  
  // const fetchApidata = async () => {
  //   try {
  //     const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
  //     const res = await web.lists
  //       .getByTitle("Expenses")
  //       .items.select(
  //         "Id",
  //         "ExpenseDate",
  //         "PaymentMode",
  //         "ExpenseType",
  //         "Amount",
  //         "AmountType",
  //         "Comment",
  //         "Modified",
  //         "Created",
  //         "Author/Id",
  //        " Author/Title",
  //         "Editor/Id",
  //         "Editor/Title"
  //       ).expand("Author,Editor")
  //       .top(4999)
  //       .get();
  
  //     const filteredData = res.filter((item: any) => item.AmountType === 'Expense');
  
  //     const formattedData = filteredData.map((item: any) => ({
  //      Id: item.Id, 
  //     ExpenseDate :item.ExpenseDate || null,
                
                 
  //      PaymentMode: item.PaymentMode || " ",
  //      ExpenseType: item.ExpenseType ||  " ",
  //      Amount: item.Amount || 0,
  //      AmountType: item.AmountType || " ",
  //      Modified: item.Modified ? moment(item.Modified).format("DD/MM/YYYY") : null,
  //      Created: item.Created ? moment(item.Created).format("DD/MM/YYYY") : null,
  //      Author: item.Author 
  //        ? { Id: item.Author.Id, Name: item.Author.Title } 
  //        : { Id: "N/A", Name: "N/A" },
  //      Editor: item.Editor 
  //        ? { Id: item.Editor.Id, Name: item.Editor.Title } 
  //        : { Id: "N/A", Name: "N/A" },
  //    }));
  //     setData(formattedData);
  //   } catch (error) {
  //     console.error("Error fetching data:", error);
  //   }
  // };

  // React.useEffect(() => {
  //   fetchApidata();
  // }, [expenseData]);

 React.useEffect(() => {
    setData(expenseData); 
  }, [expenseData]); 
  

  const handleAddTask = async () => {
    if (
      inputValue.ExpenseDate.trim() &&
      inputValue.ExpenseType.trim() &&
      inputValue.PaymentMode.trim()
    ) {
      try {
        const postData = {
          ExpenseDate: inputValue.ExpenseDate
            ? new Date(inputValue.ExpenseDate).toISOString()
            : null,
          PaymentMode: inputValue.PaymentMode || "",
          ExpenseType: inputValue.ExpenseType || "",
          Amount: inputValue.Amount || 0,
          AmountType: inputValue.AmountType,
          Comment:inputValue.Comment || "",
        };
  
        const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
        let res = await web.lists
          .getById("C72D4C84-CD85-425F-AC23-ABC284DC4406")
          .items.add(postData);
        console.log("Added item response:", res);
  
      
        // fetchApidata();
        // expenseData.fetchExpenseAPIData(currentMonth);
        fetchExpenseAPIData(selectedMonth)
        
        setInputValue({
          ExpenseDate: '',
          PaymentMode: 'UPI',
          ExpenseType: '',
          Amount: 0,
          AmountType: 'Expense',
          Comment:""
        });
  
       
        setIsPanelOpen(false);
  
       
        Swal.fire({
          // title: "Good job!",
          text: "You have successfully added items!",
          icon: "success",
        });
      } catch (error) {
        console.error("Error adding item:", error);
        Swal.fire({
          title: "Error!",
          text: "An error occurred while adding the item.",
          icon: "error",
        });
      }
    } else {
      Swal.fire({
        title: "Warning!",
        text: "Please fill out all required fields.",
        icon: "warning",
      });
    }
  };
  
  const handleInputChange = (field: keyof Expense, value: string | number) => {
    setInputValue((prev) => ({ ...prev, [field]: value }));
  };

  
  const handleEdit = (item: any) => {
    setEditId(item.Id);
    const selecteditem = item;
    setInputValue({
      ExpenseDate: item.ExpenseDate
      ? moment(item.ExpenseDate).format("YYYY-MM-DD") // Correct parsing
      : "",
        PaymentMode: selecteditem?.PaymentMode ,
        Amount: selecteditem?.Amount || 0,
        ExpenseType: selecteditem?.ExpenseType ,
        AmountType: selecteditem.AmountType || 'Expense',
        Comment:selecteditem.Comment || ""
    });
    setCreatedname(selecteditem.Author?.Name || "");
    setCreateddate(selecteditem.Created || "");
    setModifiedname(selecteditem.Editor?.Name || "");
    setModifieddate(selecteditem.Modified || "");
    setIsPanelOpen(true);
  };

  const updateDetails = async (id: number) => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      await web.lists
        .getById("C72D4C84-CD85-425F-AC23-ABC284DC4406")
        .items.getById(id)
        .update({
         ExpenseDate: inputValue.ExpenseDate
        ? new Date(inputValue.ExpenseDate).toISOString()
        : null,
            PaymentMode: inputValue?.PaymentMode || "",
            Amount: inputValue?.Amount || 0,
            ExpenseType: inputValue?.ExpenseType || "",
            AmountType: inputValue?.AmountType || "Expense",
            Comment:inputValue?.Comment || ""
            
        });

      setIsPanelOpen(false);
      setEditId(null);

      // fetchApidata();
      // expenseData.fetchExpenseAPIData(currentMonth);
      fetchExpenseAPIData(selectedMonth)
    } catch (error) {
      console.error("Error updating item:", error);
    }
  };

  const handleSaveTask = async () => {
    try {
      if (editId !== null ) {
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


  const handleDelete = async (id: number) => {
    Swal.fire({
      title: "Are you sure Delete Item",
      text: "You want to delete this data",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "delete",
    }).then(async (result) => {
      if (result.isConfirmed) {
        try {
          const web = new Web('https://smalsusinfolabs.sharepoint.com/sites/F4S');
          await web.lists
            .getById("C72D4C84-CD85-425F-AC23-ABC284DC4406")
            .items.getById(id)
            .delete()
            .then(() => {
              const remaindata = data.filter((item) => item.Id != id);
              setData(remaindata);
              // fetchApidata();
              // expenseData.fetchExpenseAPIData(currentMonth);
              fetchExpenseAPIData(selectedMonth)
  
              Swal.fire({
                title: "Deleted!",
                text: "The item has been deleted successfully.",
                icon: "success",
              });
            });
        } catch (error) {
          console.log("data in not delete");

          Swal.fire({
            title: "Error!",
            text: "Something went wrong. The item could not be deleted.",
            icon: "error",
          });
        }
      }
    });
  };

  // const openPanel = () => setIsPanelOpen(true);
  const closePanel = () => {
    setInputValue({
        ExpenseDate:"",
        PaymentMode:"",
        ExpenseType:"",
        Amount:0,
        AmountType:"Expense",
        Comment:""
    });

    setIsPanelOpen(false);
  };

  const paymentModeOptions: IDropdownOption[] = [
    { key: 'UPI', text: 'UPI' },
   { key: 'Net Banking', text: 'Net Banking' },
    { key: 'Cash', text: 'Cash' },
  ];
  const ExpenseTypeOptions: IDropdownOption[] = [
    { key: 'Electricitybill', text: 'Electricity Bill' },
   { key: 'Salary', text: 'Salary' },
    { key: 'Equipment purchases and repairs', text: 'Equipment purchases and repairs' },
    { key: 'Cleaning Staff', text: 'Cleaning Staff' },
    { key: 'Others', text: 'Others' },
  ];

  const handleNewTask = () => {
    setInputValue({
        ExpenseDate:"",
        PaymentMode:"",
        ExpenseType:"",
        Amount:0,
        AmountType:"Expense",
        Comment:"",
    });
    setEditId(null);
    setIsPanelOpen(true);
  };

  const onRenderFooterContent = React.useCallback(() => {
    const formatDate = (date:any) => {
      if (!date) return "N/A";
      return moment(date).format("DD/MM/YYYY"); 
    };
  
    return (
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
        {editId != null && (
          <div>
            <div>
              Created{" "}
              <span style={{ color: "skyblue", fontSize: "10px" }}>
                {formatDate(Createddate)}
              </span>{" "}
              by{" "}
              <span style={{ color: "skyblue", fontSize: "10px" }}>
                {Createdname || "N/A"}
              </span>
            </div>
            <div>
              Last modified{" "}
              <span style={{ color: "skyblue", fontSize: "10px" }}>
                {formatDate(Modifiedate)}
              </span>{" "}
              by{" "}
              <span style={{ color: "skyblue", fontSize: "10px" }}>
                {Modifiedname || "N/A"}
              </span>
            </div>
          </div>
        )}
  
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {editId != null && (
            <a
              href={`https://smalsusinfolabs.sharepoint.com/sites/IITIQ/Lists/StaffMembers/EditForm.aspx?ID=${editId}`}
              style={{ textDecoration: "none", color: "skyblue", fontSize: "14px" }}
              target="_blank"
              rel="noopener noreferrer"
            >
              Open out-of-the-box form
            </a>
          )}
          <PrimaryButton onClick={handleSaveTask}>Save</PrimaryButton>
          <DefaultButton onClick={() => closePanel()}>Cancel</DefaultButton>
        </div>
      </div>
    );
  }, [handleSaveTask, closePanel, Createddate, Createdname, Modifiedate, Modifiedname]);
  

  return (
    <div style={{ backgroundColor: 'white',marginTop:'20px'}}>
       <div className="d-flex justify-content-between align-items-center mb-2">
       <div className="flex-grow-1 text-center">
       <h5 className="mb-0">Expense Table</h5>
       </div>
      <button
       className="btn px-3 py-2 rounded-2 btn btn-primary"
       onClick={handleNewTask}
       >
    Add
  </button>
    </div>
    <Panel
     isOpen={isPanelOpen}
     onDismiss={closePanel}
    headerText="Add Expense"
    onRenderFooterContent={onRenderFooterContent}
    closeButtonAriaLabel="Close"
    isFooterAtBottom={true}
    type = {PanelType.medium}
   >
  <div className='row'>
  <div className="col-lg-6">
  <div className="form-group m-2">
      <h6>Expense Date</h6>
      <TextField
        placeholder="Enter Expense"
         type='date'
        value={inputValue.ExpenseDate}
        onChange={(e) => handleInputChange("ExpenseDate", e.currentTarget.value)}
        style={{ width: '100%' }}
        required
        errorMessage={!inputValue?.ExpenseDate ? "expense date  is required" : ""}
   
      />
    </div>
    </div>
    <div className="col-lg-6">
  <div className="form-group m-2">
      <h6>Payment Mode</h6>
      <Dropdown
        selectedKey={inputValue.PaymentMode}
        options={paymentModeOptions}
        onChange={(e, option) => handleInputChange("PaymentMode", option?.key || "")}
        styles={{ dropdown: { width: '100%' } }}
      />
    </div>
    </div>
    <div className="col-lg-6">
    <div className="form-group m-2">
      <h6>Amount</h6>
      <TextField
        placeholder="Enter Amount"
        type="number"
        value={inputValue.Amount !== undefined ? inputValue.Amount.toString() : ""}
        onChange={(e) => handleInputChange("Amount", Number(e.currentTarget.value))}
        style={{ width: '100%' }}
      />
    </div>
    </div>
    <div className="col-lg-6">
    <div className="form-group m-2">
      <h6>Amount Type</h6>
      <TextField
        disabled
        value={inputValue.AmountType}
        style={{ width: '100%' }}
      />
    </div>
    </div>
    <div className="col-lg-6">
    <div className="form-group m-2">
      <h6>Expense Type</h6>
      <Dropdown
        selectedKey={inputValue.ExpenseType}
        options={ExpenseTypeOptions}
        onChange={(e, option) => handleInputChange("ExpenseType", option?.key || "")}
        styles={{ dropdown: { width: '100%' } }}
      />
    </div>
    </div>
    <div className="col-lg-6">
    <div className="form-group m-2">
      <h6>Comment</h6>
      <textarea
        value={inputValue.Comment}
        onChange={(e) => handleInputChange("Comment", e.currentTarget.value)}
        style={{ width: '100%' }}
      />
    </div>
    </div>
    {/* <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
      <PrimaryButton styles={{ root: { marginRight: 8 } }} onClick={handleAddTask}>
        Save
      </PrimaryButton>
      <DefaultButton onClick={closePanel}>Cancel</DefaultButton>
    </div> */}
  </div>
      </Panel>



     
      <div style={{ maxHeight: "500px", overflowY: "auto", }}>
      <table
        style={{
          borderCollapse: 'collapse',
          width: '100%',
           border: '1px solid #ddd',
          maxHeight: "500px",
          overflowY: "auto" 
         
        }}
      >
        <thead>
          <tr style={{ backgroundColor: '#f4f4f4',textAlign:"center" }}>
            <th style={{ padding: '8px', border: '1px solid #ddd',textAlign:"center" }}>Expense Date</th>
            <th style={{ padding: '8px', border: '1px solid #ddd',textAlign:"center" }}>Payment Type</th>
            <th style={{ padding: '8px', border: '1px solid #ddd',textAlign:"center" }}>Expense Type</th>
            <th style={{ padding: '8px', border: '1px solid #ddd',textAlign:"center" }}>Description</th>
            <th style={{ padding: '8px', border: '1px solid #ddd',textAlign:"center" }}>Amount</th>
          </tr>
        </thead>
        <tbody>
          {data.length > 0 ? (
            expenseData.map((row:any, index:any) => (
              <tr key={index} style={{ textAlign: 'center' }}>
               <td style={{ padding: '8px', border: '1px solid #ddd' ,textAlign: 'center'}}>
               {moment(row.ExpenseDate).format('DD/MM/YYYY')}
            </td>
                <td style={{ padding: '8px', border: '1px solid #ddd',textAlign: 'center' }}>
                  {row.PaymentMode}
                </td>
                <td style={{ padding: '8px', border: '1px solid #ddd',textAlign: 'center' }}>
                  {row.ExpenseType}
                </td>
                <td style={{ padding: '8px', border: '1px solid #ddd',textAlign: 'center' }}>
                  {row.Comment}
                </td>
                <td style={{ padding: '8px', border: '1px solid #ddd',textAlign: 'center' }}>
  ₹{(row?.Amount ?? 0).toFixed(2)}
</td>



                <td style={{ padding: "10px", border: "1px solid #ddd",textAlign: 'center' }}>
         <div className="d-flex">
            <span style={{ color: "green", cursor: "pointer", padding: "8px" }} onClick={() => handleEdit(row)}>
              <FaEdit />
            </span>
            <span style={{ color: "#ff0000", cursor: "pointer", padding: "8px" }} onClick={() => handleDelete(row.Id)}>
              <MdDelete />
            </span>
         </div>
        </td>
              </tr>
            ))
          ) : (
            <tr>
              <td
                colSpan={4}
                style={{
                  padding: '8px',
                  textAlign: 'center',
                  border: '1px solid #ddd',
                }}
              >
               
              </td>
            </tr>
          )}
        </tbody>
      </table>
      </div>
    </div>
  );
};

export default Expensetable;


