import * as React from 'react';
import  {useState,useCallback,useEffect} from 'react';
// import { Web } from 'sp-pnp-js';
import { DefaultButton, PrimaryButton, TextField , Dropdown, IDropdownOption} from '@fluentui/react';
import { Panel ,PanelType  } from '@fluentui/react/lib/Panel';
import moment from 'moment';


interface PaymentHistoryProps {
   userId: number | null;
  onPaymentHistoryUpdate: (data: any[]) => void;
  paymentHistoryData:any[];
}
interface PaymentItem {
  Id: number;
  Amount: string;
  Received: string;
  Paymentmode: string;
  Date: string;
  status: string;
  comments?: string;
}



const Paymenthistory:React.FC<PaymentHistoryProps> = ({ userId,onPaymentHistoryUpdate,paymentHistoryData}) => {
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [history, setHistory] = useState<PaymentItem[]>([]);
  const [inputValue, setInputValue] = useState({
    Amount: "",
    Received:"",
    Paymentmode: "",
    Date: "",
    status: "",
    comments: "",
  });
  const [editPaymentId, setEditPaymentId] = useState(null);


console.log("paymentHistoryData", paymentHistoryData);

const handleAddPayment = () => {
  const formattedDate = inputValue.Date
    ? moment(inputValue.Date).format("DD/MM/YYYY")
    : "";

  const newPayment = {
    Id: editPaymentId || Date.now(),
    Amount: inputValue.Amount || "",
    Received:inputValue.Received || "",
    Paymentmode: inputValue.Paymentmode?.trim() || "",
    Date: formattedDate,
    status: inputValue.status?.trim() || "",
    comments: inputValue.comments || "",
  };

  let updatedHistory;

  if (editPaymentId) {
    updatedHistory = history.map((payment) =>
      payment.Id === editPaymentId ? newPayment : payment
    );
  } else {
    updatedHistory = [...history, newPayment];
  }

  setHistory(updatedHistory);
  onPaymentHistoryUpdate(updatedHistory);

  // Reset form
  setInputValue({
    Amount: "",
    Received:"",
    Paymentmode: "",
    Date: "",
    status: "",
    comments: "",
  });
  setEditPaymentId(null);
  setIsPanelOpen(false);
};


const handleEditPayment = (payment:any) => {
  const parsedDate = moment(payment.Date, ["DD/MM/YYYY", "YYYY-MM-DD"], true);
  setInputValue({
    Amount: payment?.Amount,
    Received:payment?.Received,
    Paymentmode: payment?.Paymentmode,
    Date: parsedDate.isValid() ? parsedDate.format("YYYY-MM-DD") : "",
    status: payment?.status,
    comments: payment?.comments || "",
  });
  setEditPaymentId(payment.Id);
  setIsPanelOpen(true);
};




useEffect(() => {
    if (userId !== null) {
      setHistory(paymentHistoryData || []);
    } else {
      setHistory([]);
    }
  }, [userId, paymentHistoryData]);


const handleNewPayment = () => {
     setInputValue({
        Amount: "",
        Received:"",
        Paymentmode: "",
        Date: "",
        status: "",
        comments: "",
      });
    setIsPanelOpen(true);
  };



  const handleClosePanel = () => {
    setIsPanelOpen(false);
  };


  const handleInputChange = (key: string, value: any) => {
    setInputValue((prev) => ({
      ...prev,
      [key]: value,
    }));
  };

 const validPaymentData = history?.filter((item) =>
  !!(
    item.Amount?.trim() ||
    item.Received?.trim()||
    item.Paymentmode?.trim() ||
    item.Date?.trim() ||
    item.status?.trim() ||
    item.comments?.trim()
  )
);



  const onRenderFooterContent = useCallback(
    () => (
      <div>
        <PrimaryButton onClick={handleAddPayment }>Save</PrimaryButton>
        <DefaultButton onClick={handleClosePanel}>Cancel</DefaultButton>
      </div>
    ),
    [handleAddPayment,handleClosePanel]
  );

  const paymentModeOptions: IDropdownOption[] = [
    { key: "UPI", text: "UPI" },
    { key: "Net Banking", text: "Net Banking" },
    { key: "Cash", text: "Cash" },
  ];
   const statusOptions:IDropdownOption[] = [
      { key: "Paid", text: "Paid" },
      { key: "Pending", text: "Pending" },
    ];
    
  return (
    <>
    <div className="container mt-4">
      <div className="border rounded shadow-sm bg-white p-3">
  <div className="d-flex justify-content-between align-items-center mb-3">
     <h6 className="text-center w-100 mb-0 fw-medium text-secondary" style={{ fontSize: "1.1rem" }}>
        Payment History Table
      </h6>
    <button
        className="btn btn-success btn-sm "
        onClick={handleNewPayment}
      >
         Add
      </button>
  </div>

 <Panel
  isOpen={isPanelOpen}
  onDismiss={handleClosePanel}
  headerText="Payment details"
  closeButtonAriaLabel="Close"
  onRenderFooterContent={onRenderFooterContent}
  isFooterAtBottom={true}
  type={PanelType.medium}
>
  <div className="container">
    <div className="row">
      <div className="col-12 form-group mb-3">
        <h6>Payment Mode</h6>
        <Dropdown
          placeholder="Select Payment Mode"
          options={paymentModeOptions}
          selectedKey={inputValue.Paymentmode}
          onChange={(e, option) => handleInputChange("Paymentmode", option?.key)}
        />
      </div>

      <div className="col-12 form-group mb-3">
        <h6>Amount</h6>
        <TextField
          placeholder="Enter Amount"
          type="number"
          value={inputValue.Amount !== undefined ? inputValue.Amount?.toString() : ""}
          onChange={(e) => {
            const value = Number(e.currentTarget.value);
            handleInputChange("Amount", value);
          }}
        />
      </div>
      <div className="col-12 form-group mb-3">
        <h6>Amt. Received</h6>
        <TextField
          placeholder="Enter received amount"
          type="number"
          value={inputValue.Received !== undefined ? inputValue.Received?.toString() : ""}
          onChange={(e) => {
            const value = Number(e.currentTarget.value);
            handleInputChange("Received", value);
          }}
        />
      </div>

      <div className="col-12 form-group mb-3">
        <h6>Date</h6>
        <TextField
          placeholder="Enter Date"
          type="date"
          value={inputValue.Date}
          onChange={(e) => handleInputChange("Date", e.currentTarget.value)}
        />
      </div>

      <div className="col-12 form-group mb-3">
        <h6>Status</h6>
        <Dropdown
          placeholder="Select Status"
          options={statusOptions}
          selectedKey={inputValue.status}
          onChange={(e, option) => handleInputChange("status", option?.key)}
        />
      </div>

      <div className="col-12 form-group mb-3">
        <h6>Comments</h6>
        <TextField
          multiline
          rows={3}
          placeholder="Enter comment"
          value={inputValue.comments}
          onChange={(e) => handleInputChange("comments", e.currentTarget.value)}
        />
      </div>
    </div>
  </div>
</Panel>

 <div className="p-4 rounded shadow-lg border border-2">
  <table className="table table-bordered table-hover table-striped mb-0">
    <thead className="table-light text-center">
      <tr>
        <th>Amount</th>
        <th>Payment Mode</th>
        <th>Date</th>
        <th>Status</th>
        <th>Comments</th>
        <th>Actions</th>
      </tr>
    </thead>
    <tbody>
      {validPaymentData && validPaymentData.length > 0 ? (
        validPaymentData.map((item) => (
          <tr key={item.Id}>
            <td>{item.Amount}</td>
            <td>{item.Paymentmode}</td>
            <td>{item.Date}</td>
            <td>{item.status}</td>
            <td>{item.comments}</td>
            <td className="text-center">
              <button
                className="btn btn-sm btn-primary"
                onClick={() => handleEditPayment(item)}
              >
                Edit
              </button>
            </td>
          </tr>
        ))
      ) : (
        <tr>
          <td colSpan={6} className="text-center text-muted py-3">
            <em>No payment history found.</em>
          </td>
        </tr>
      )}
    </tbody>
  </table>
</div>

  </div>
</div>

    </>
  );
};

export default Paymenthistory;


