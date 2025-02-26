import * as React from 'react';
import { Web } from 'sp-pnp-js';
import { DefaultButton, PrimaryButton, TextField , Dropdown, IDropdownOption} from '@fluentui/react';
import { Panel,PanelType   } from '@fluentui/react/lib/Panel';
import moment from 'moment';
import Swal from 'sweetalert2';
import "bootstrap/dist/css/bootstrap.min.css";

type Sale = {
 ExpenseDate: string; 
  PaymentMode: string;
  SalesType: string;
  Amount: number;
  AmountType: string;
};



const Salestable = ( salesData:any) => {
  let currentMonth =salesData?.selectedMonth
  const [data, setData] = React.useState<Sale[]>([]);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [inputValue, setInputValue] = React.useState({

    ExpenseDate:'',
    PaymentMode:'',
    SalesType:'',
    Amount:0,
    AmountType:'Sales',
   
  })

 
  const fetchAPIData = async () => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const res = await web.lists
        .getByTitle("Expenses")
        .items.select(
          "ExpenseDate",
          "PaymentMode",
          "SalesType",
          "Amount",
          "AmountType"
        )
        .top(4999)
        .get();

        const filteredData = res.filter((item: any) => item.AmountType === 'Sales');

        const formattedData = filteredData.map((item: any) => ({
          ExpenseDate: item.ExpenseDate
            ? moment(item.ExpenseDate).format('DD/MM/YYYY') 
            : null,
          PaymentMode: item.PaymentMode || 'N/A',
          SalesType: item.SalesType || 'N/A',
          Amount: item.Amount || 0,
        }));

      setData(formattedData);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  React.useEffect(() => {
    fetchAPIData();
  }, [salesData]);

  const handleAddTask = async () => {
    if (
      inputValue.ExpenseDate.trim() &&
      inputValue.SalesType.trim() &&
      inputValue.PaymentMode.trim()
    ) {
      try {
        const postData = {
          ExpenseDate: inputValue.ExpenseDate
            ? new Date(inputValue.ExpenseDate).toISOString()
            : null,
          PaymentMode: inputValue.PaymentMode,
          SalesType: inputValue.SalesType,
          Amount: inputValue.Amount,
          AmountType: inputValue.AmountType,
        };
  
        const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
        let res = await web.lists
          .getById("C72D4C84-CD85-425F-AC23-ABC284DC4406")
          .items.add(postData);
        console.log("Added item response:", res);
  
      
        fetchAPIData();
        salesData.fetchExpenseAPIData(currentMonth);
  
      
        setInputValue({
          ExpenseDate: '',
          PaymentMode: '',
          SalesType: '',
          Amount: 0,
          AmountType: 'Sales',
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
  


  const handleInputChange = (field: keyof  Sale, value: string | number) => {
    setInputValue((prev) => ({ ...prev, [field]: value }));
  };


  const openPanel = () => setIsPanelOpen(true);
  const closePanel = () => setIsPanelOpen(false);


  const paymentModeOptions: IDropdownOption[] = [
    { key: 'UPI', text: 'UPI' },
   { key: 'Net Banking', text: 'Net Banking' },
    { key: 'Cash', text: 'Cash' },
  ];
  const SalesTypeOptions: IDropdownOption[] = [
    { key: 'Proteins', text: 'Proteins' },
   { key: 'Creatine', text: 'creatine' },
    { key: 'Gym accessories', text: 'gym accessories' },
    { key: 'others', text: 'Others' },
    
  ];

  return (
    <div style={{ backgroundColor: "white",marginTop:"20px" }}>
       <div className="d-flex justify-content-between align-items-center mb-2">
       <div className="flex-grow-1 text-center">
       <h5 className="mb-0">Sales Table</h5>
       </div>
      <button
       className="btn px-3 py-2 rounded-2 btn btn-primary"
       onClick={openPanel}
       >
    Add
  </button>
</div>

   <Panel
     isOpen={isPanelOpen}
     onDismiss={closePanel}
    headerText="Add Sales"
    closeButtonAriaLabel="Close"
    isFooterAtBottom={true}
    type = {PanelType.medium}
   >

  <div className='row'>
  <div className="col-lg-6">
  <div className="form-group m-2">
      <h6>Sales Date</h6>
      <TextField
        placeholder="Enter Date"
         type='date'
        value={inputValue.ExpenseDate}
        onChange={(e) => handleInputChange("ExpenseDate", e.currentTarget.value)}
        style={{ width: '100%' }}
        required
        errorMessage={!inputValue?.ExpenseDate ? "expense date  is required" : ""}
  
      />
    </div>
    </div>
    <div className='col-lg-6'>
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
    <div className='col-lg-6'>
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
    <div className='col-lg-6'>
    <div className="form-group m-2">
      <h6>Amount Type</h6>
      <TextField
        disabled
        value={inputValue.AmountType}
        style={{ width: '100%' }}
      />
    </div>
    </div>
    <div className='col-lg-6'>
    <div className="form-group m-2">
      <h6>Sales Type</h6>
      <Dropdown
        selectedKey={inputValue.SalesType}
        options={SalesTypeOptions}
        onChange={(e, option) => handleInputChange("SalesType", option?.key || "")}
        styles={{ dropdown: { width: '100%' } }}
      />
    </div>
    </div>
    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
      <PrimaryButton styles={{ root: { marginRight: 8 } }} onClick={handleAddTask}>
        Save
      </PrimaryButton>
      <DefaultButton onClick={closePanel}>Cancel</DefaultButton>
    </div>
  </div>
</Panel>

     
      <table
        style={{
          borderCollapse: "collapse",
          width: "100%",
          border: "1px solid #ddd",
         
        }}
      >
        <thead>
          <tr style={{ backgroundColor: "#f4f4f4" }}>
            <th style={{ padding: "8px", border: "1px solid #ddd" }}>Sales Date</th>
            <th style={{ padding: "8px", border: "1px solid #ddd" }}>Payment Mode</th>
            <th style={{ padding: "8px", border: "1px solid #ddd" }}>Description</th>
            <th style={{ padding: "8px", border: "1px solid #ddd" }}>Amount</th>
          </tr>
        </thead>
        <tbody>
          {data.length > 0 ? (
          salesData?.salesData?.map((row:any, index:any)=> (
              <tr key={index} style={{ textAlign: "center" }}>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                {moment(row.ExpenseDate).format('DD/MM/YYYY')}
                </td>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                  {row.PaymentMode}
                </td>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                  {row.SalesType}
                </td>
                <td style={{ padding: "8px", border: "1px solid #ddd" }}>
                  {row.AmountType === "Sales"
                    ? `${row.Amount} `
                    : row.Amount}
                </td>
              </tr>
            ))
          ) : (
            <tr>
              <td
                colSpan={4}
                style={{
                  padding: "8px",
                  textAlign: "center",
                  border: "1px solid #ddd",
                }}
              >
               
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};

export default Salestable;
