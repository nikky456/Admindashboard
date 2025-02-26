import * as React from "react";
import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Dropdown,
  IDropdownOption,
} from "@fluentui/react";
import { Panel, PanelType } from "@fluentui/react/lib/Panel";
import { Web } from "sp-pnp-js";
import moment from "moment";

// Define interface for props if needed
interface PanelProps {
  selectedResults: any;
  onClose: () => void;
}
// Define Gender and Membership Plan options
const genderOptions: IDropdownOption[] = [
  { key: "Male", text: "Male" },
  { key: "Female", text: "Female" },
  { key: "Other", text: "Other" },
];

const paymentModeOptions: IDropdownOption[] = [
  { key: "UPI", text: "UPI" },
  { key: "Cash", text: "Cash" },
  { key: "Card", text: "Card" },
];

const MembershipPlansOptions: IDropdownOption[] = [
  { key: "PlanA", text: "Plan A" },
  { key: "PlanB", text: "Plan B" },
  { key: "PlanC", text: "Plan C" },
];

const SearchPanel: React.FC<PanelProps> = ({ selectedResults, onClose }) => {
  const [editId, setEditId] = React.useState<number | null>(null);
  const [data, setData] = React.useState<any[]>([]);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [inputValue, setInputValue] = React.useState({
    Id: 0,
    FirstName: "",
    LastName: "",
    FullName: "",
    Age: 0,
    Gender: "Male",
    Email: "",
    AadhaarNumber: 0,
    CellPhone: "",
    WorkAddress: "",
    JoiningDate: "",
    EndDate: "",
    PaymentMode: "UPI",
    BillAmount: 0,
    AmountReceived: 0,
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
    MembershipNo: 0,
    MembershipPlan: "",
    BMIDetails: [
      {
        date: "",
        name: "",
        contact: "",
        height: "",
        weight: "",
        age: "",
        vFat: "", // Visceral Fat
        kcal: "", // Calorie Intake
        bodyAge: "",
        fat: "",
        bmi: "",
        bmiCategory: "",
        subWholeBody: "",
        subTrunk: "",
        subArms: "",
        subLegs: "",
        sMusclesWholeBody: "",
        sMusclesTrunk: "",
        sMusclesArms: "",
        sMusclesLegs: "",
      },
    ],
    DietChart: [
      {
        mealTime: "",
        mealDescription: "",
        caloriesPortion: "",
        timeSlot: "",
        dayOfWeek: "",
      },
    ],
    ProgressReport: [
      {
        date: "",
        type: "BMI",
        value: "",
        status: "",
      },
    ],
  });

  // const [isEditMode, setIsEditMode] = React.useState(false);
  const fetchAPIData = async () => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const res = await web.lists
        .getByTitle("Clients")
        .items.select(
          "Id",
          "FirstName",
          "LastName",
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
          "BMIDetails",
          "Progress Report",
          "Diet Chart"
        )
        .top(4999)
        .get();

      setData(res);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  console.log("Inside Search panel");
  console.log(data);

  React.useEffect(() => {
    fetchAPIData();
  }, []);

  React.useEffect(() => {
    if (selectedResults) {
      setInputValue({
        ...inputValue,
        ...selectedResults,
        MedicalDetails: selectedResults.MedicalDetails || [
          {
            BloodGroup: "",
            BP: 0,
            HeartBeat: 0,
            Sugar: 0,
            Others: "",
          },
        ],
      });
      setIsPanelOpen(true); // Open the panel
    }
  }, [selectedResults]);

  React.useEffect(() => {
    console.log("Selected Client in Panel:", selectedResults);
  }, [selectedResults]);

  const updateDetails = async (id: number, inputValue: any) => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      await web.lists
        .getById("3A9C0B25-B14D-4277-99CB-D63FCFF5FD3F")
        .items.getById(id)
        .update({
          FirstName: inputValue.FirstName,
          LastName: inputValue.LastName,
          Title: inputValue.Title,
          FullName: inputValue.FullName,
          Email: inputValue.Email,
          Age: inputValue.Age,
          Gender: inputValue.Gender,
          AadhaarNumber: inputValue.AadhaarNumber,
          CellPhone: inputValue.CellPhone,
          WorkAddress: inputValue.WorkAddress,
          JoiningDate: inputValue.JoiningDate
            ? moment(inputValue.JoiningDate, "YYYY-MM-DD").format("MM-DD-YYYY")
            : null,
          EndDate: inputValue.EndDate
            ? moment(inputValue.EndDate, "YYYY-MM-DD").format("MM-DD-YYYY")
            : null,
          PaymentMode: inputValue.PaymentMode,
          AmountReceived: inputValue.AmountReceived,
          MedicalDetails: JSON.stringify(inputValue.MedicalDetails),
          BMIDate: inputValue.BMIDate
            ? moment(inputValue.BMIDate, "YYYY-MM-DD").format("MM-DD-YYYY")
            : null,
          NextBMIDueDate: inputValue.NextBMIDueDate
            ? moment(inputValue.NextBMIDueDate, "YYYY-MM-DD").format(
                "MM-DD-YYYY"
              )
            : null,
          MembershipNo: inputValue.MembershipNo,
          MembershipPlan: inputValue.MembershipPlan,
        });
      fetchAPIData();
    } catch (error) {
      console.error("Error updating item:", error);
    }
  };

  // const handleEditTask = (rowData: any) => {
  //   console.log("Row data on edit:", rowData); // Check row data
  //   setEditId(rowData.Id);
  //   const parsedRowData = {
  //     ...rowData,
  //     MedicalDetails: rowData.MedicalDetails
  //       ? JSON.parse(rowData.MedicalDetails)
  //       : [
  //           {
  //             BloodGroup: "",
  //             BP: 0,
  //             HeartBeat: 0,
  //             Sugar: 0,
  //             Others: "",
  //           },
  //         ],
  //   };

  //   console.log("Parsed data:", parsedRowData); // Check parsed data
  //   setInputValue(parsedRowData);
  //   setIsPanelOpen(true);
  // };

  // const handleSaveTask = async () => {
  //   if (editId !== null) {
  //     await updateDetails(editId, inputValue);
  //   } else {
  //     console.log("Add functionality to handle new tasks.");
  //   }
  //   setEditId(null);
  //   setIsPanelOpen(false);
  // };

  // const handleClick = async (item: any) => {
  //   await handleEditTask(item);
  //   await handleSaveTask();
  // };

  const handleSaveTask = async () => {
    if (!inputValue) {
      console.error("No input value provided");
      return;
    }

    // Edit logic (process data as in handleEditTask)
    const parsedInputValue = {
      ...inputValue,
      MedicalDetails: inputValue.MedicalDetails
        ? JSON.stringify(inputValue.MedicalDetails)
        : [
            {
              BloodGroup: "",
              BP: 0,
              HeartBeat: 0,
              Sugar: 0,
              Others: "",
            },
          ],
    };

    console.log("Processed data for saving:", parsedInputValue);

    // Save to backend
    if (editId !== null) {
      try {
        await updateDetails(editId, parsedInputValue);
        console.log("Data successfully saved!");
      } catch (error) {
        console.error("Error saving data:", error);
      }
    } else {
      console.log("Add functionality to handle new tasks.");
    }

    // Reset states after saving
    setEditId(null);
    setIsPanelOpen(false);
  };

  const handleClosePanel = () => {
    setInputValue({
      Id: 0,
      FirstName: "",
      LastName: "",
      FullName: "",
      Age: 0,
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
      MembershipNo: 0,
      MembershipPlan: "",
      BMIDetails: [
        {
          date: "",
          name: "",
          contact: "",
          height: "",
          weight: "",
          age: "",
          vFat: "", // Visceral Fat
          kcal: "", // Calorie Intake
          bodyAge: "",
          fat: "",
          bmi: "",
          bmiCategory: "",
          subWholeBody: "",
          subTrunk: "",
          subArms: "",
          subLegs: "",
          sMusclesWholeBody: "",
          sMusclesTrunk: "",
          sMusclesArms: "",
          sMusclesLegs: "",
        },
      ],
      DietChart: [
        {
          mealTime: "",
          mealDescription: "",
          caloriesPortion: "",
          timeSlot: "",
          dayOfWeek: "",
        },
      ],
      ProgressReport: [
        {
          date: "",
          type: "BMI",
          value: "",
          status: "",
        },
      ],
    });

    setIsPanelOpen(false);
    onClose();
  };

  const handleInputChange = (
    field: keyof typeof inputValue,
    value: string | number
  ) => {
    setInputValue((prev) => ({ ...prev, [field]: value }));
    console.log(inputValue);
  };

  const handleMedicalDetailsChange = (field: string, value: any) => {
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

  return (
    <div>
      <Panel
        isOpen={isPanelOpen}
        onDismiss={handleClosePanel}
        headerText="Form Details"
        closeButtonAriaLabel="Close"
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
              <h6>First Name</h6>
              <TextField
                placeholder="Enter First Name"
                value={inputValue.FirstName}
                onChange={(e) =>
                  handleInputChange("FirstName", e.currentTarget.value)
                }
              />
            </div>
            <div className="col">
              <h6>Last Name</h6>
              <TextField
                placeholder="Enter Last Name"
                value={inputValue.LastName}
                onChange={(e) =>
                  handleInputChange("LastName", e.currentTarget.value)
                }
              />
            </div>
            <div className="col">
              <h6>Full Name</h6>
              <TextField
                placeholder="Enter Full Name"
                value={inputValue.FullName}
                onChange={(e) =>
                  handleInputChange("FullName", e.currentTarget.value)
                }
                // disabled={!isEditMode}
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
                  handleInputChange(
                    "Age",
                    isNaN(parsedValue) ? "" : parsedValue
                  );
                }}
                // disabled={!isEditMode}
              />
            </div>

            <div className="col">
              <h6>Gender</h6>
              <Dropdown
                selectedKey={inputValue.Gender}
                options={genderOptions}
                onChange={(e, option) =>
                  handleInputChange("Gender", option?.key || "")
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>Email</h6>
              <TextField
                placeholder="Enter Email"
                value={inputValue.Email}
                onChange={(e) =>
                  handleInputChange("Email", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
          </div>

          {/* Row 3 */}
          <div className="row align-items-center mb-4">
            <div className="col">
              <h6>Aadhaar Number</h6>
              <TextField
                placeholder="Enter Aadhaar Number"
                type="number"
                value={
                  inputValue.AadhaarNumber !== undefined
                    ? inputValue.AadhaarNumber?.toString()
                    : ""
                }
                onChange={(e) =>
                  handleInputChange(
                    "AadhaarNumber",
                    Number(e.currentTarget.value)
                  )
                }
                // disabled={!isEditMode}
              />
            </div>

            <div className="col">
              <h6>Cell Phone</h6>
              <TextField
                placeholder="Enter Cell Phone"
                value={inputValue.CellPhone}
                onChange={(e) =>
                  handleInputChange("CellPhone", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>Work Address</h6>
              <TextField
                placeholder="Enter Work Address"
                value={inputValue.WorkAddress}
                onChange={(e) =>
                  handleInputChange("WorkAddress", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
          </div>

          {/* Row 4 */}
          <div className="row align-items-center mb-4">
            <div className="col">
              <h6>Joining Date</h6>
              <TextField
                placeholder="Enter Joining Date"
                type="date"
                value={inputValue.JoiningDate}
                onChange={(e) =>
                  handleInputChange("JoiningDate", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>End Date</h6>
              <TextField
                placeholder="Enter End Date"
                type="date"
                value={inputValue.EndDate}
                onChange={(e) =>
                  handleInputChange("EndDate", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>Payment Mode</h6>
              <Dropdown
                selectedKey={inputValue.PaymentMode}
                options={paymentModeOptions}
                onChange={(e, option) =>
                  handleInputChange("PaymentMode", option?.key || "")
                }
                // disabled={!isEditMode}
              />
            </div>
          </div>
          {/* Row 5 */}
          <div className="row align-items-center mb-4">
            <div className="col-12">
              <h6>Medical Details</h6>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(5, 1fr)",
                  gap: "16px",
                  width: "100%",
                }}
              >
                <Dropdown
                  placeholder="Select Blood Group"
                  selectedKey={inputValue.MedicalDetails[0].BloodGroup}
                  options={[
                    { key: "A+", text: "A+" },
                    { key: "A-", text: "A-" },
                    { key: "B+", text: "B+" },
                    { key: "B-", text: "B-" },
                    { key: "O+", text: "O+" },
                    { key: "O-", text: "O-" },
                    { key: "AB+", text: "AB+" },
                    { key: "AB-", text: "AB-" },
                  ]}
                  onChange={(e, option) =>
                    handleMedicalDetailsChange("BloodGroup", option?.key || "")
                  }
                  // disabled={!isEditMode}
                  styles={{ dropdown: { width: "100%" } }}
                />
                <TextField
                  placeholder=" Enter BP"
                  type="number"
                  value={
                    inputValue.MedicalDetails[0]?.BP
                      ? inputValue.MedicalDetails[0].BP?.toString()
                      : ""
                  }
                  onChange={(e) =>
                    handleMedicalDetailsChange(
                      "BP",
                      Number(e.currentTarget.value) || 0
                    )
                  }
                  // disabled={!isEditMode}
                  styles={{ root: { width: "100%" } }}
                />
                <TextField
                  placeholder=" Enter Heart Beat"
                  type="number"
                  value={
                    inputValue.MedicalDetails[0]?.HeartBeat
                      ? inputValue.MedicalDetails[0].HeartBeat?.toString()
                      : ""
                  }
                  onChange={(e) =>
                    handleMedicalDetailsChange(
                      "HeartBeat",
                      Number(e.currentTarget.value) || 0
                    )
                  }
                  // disabled={!isEditMode}
                  styles={{ root: { width: "100%" } }}
                />
                <TextField
                  placeholder="Enter Sugar"
                  type="number"
                  value={
                    inputValue.MedicalDetails[0]?.Sugar
                      ? inputValue.MedicalDetails[0].Sugar?.toString()
                      : ""
                  }
                  onChange={(e) =>
                    handleMedicalDetailsChange(
                      "Sugar",
                      Number(e.currentTarget.value) || 0
                    )
                  }
                  // disabled={!isEditMode}
                  styles={{ root: { width: "100%" } }}
                />
                <TextField
                  placeholder="Others"
                  value={inputValue.MedicalDetails[0].Others}
                  onChange={(e) =>
                    handleMedicalDetailsChange("Others", e.currentTarget.value)
                  }
                  // disabled={!isEditMode}
                  styles={{ root: { width: "100%" } }}
                />
              </div>
            </div>
          </div>
          {/* row 6 */}
          <div className="row align-items-center mb-4">
            <div className="col">
              <h6>Bill Amount</h6>
              <TextField
                placeholder="Enter Bill Amount"
                type="number"
                value={
                  inputValue.BillAmount !== undefined
                    ? inputValue.BillAmount?.toString()
                    : ""
                }
                onChange={(e) =>
                  handleInputChange("BillAmount", Number(e.currentTarget.value))
                }
                // disabled={!isEditMode}
              />
            </div>

            <div className="col">
              <h6>Amount Received</h6>
              <TextField
                placeholder="Enter Amount Received"
                type="number"
                value={
                  inputValue.AmountReceived !== undefined
                    ? inputValue.AmountReceived?.toString()
                    : ""
                }
                onChange={(e) =>
                  handleInputChange(
                    "AmountReceived",
                    Number(e.currentTarget.value)
                  )
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>BMI date</h6>
              <TextField
                placeholder="Enter BMIDate"
                type="date"
                value={inputValue.BMIDate}
                onChange={(e) =>
                  handleInputChange("BMIDate", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
          </div>
          {/* row-7 */}
          <div className="row align-items-center mb-4">
            <div className="col">
              <h6>NextBMIDueDate</h6>
              <TextField
                placeholder="Enter NextBMIDueDate"
                type="date"
                value={inputValue.NextBMIDueDate}
                onChange={(e) =>
                  handleInputChange("NextBMIDueDate", e.currentTarget.value)
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6> MembershipNo</h6>
              <TextField
                placeholder="Enter  MembershipNo"
                type="number"
                value={
                  inputValue.MembershipNo !== undefined
                    ? inputValue.MembershipNo?.toString()
                    : ""
                }
                onChange={(e) =>
                  handleInputChange(
                    "MembershipNo",
                    Number(e.currentTarget.value)
                  )
                }
                // disabled={!isEditMode}
              />
            </div>
            <div className="col">
              <h6>MembershipPlan</h6>
              <Dropdown
                selectedKey={inputValue.MembershipPlan}
                options={MembershipPlansOptions}
                onChange={(e, option) =>
                  handleInputChange("MembershipPlan", option?.key || "")
                }
                // disabled={!isEditMode}
              />
            </div>
          </div>
        </div>

        <div className="d-flex justify-content-end">
          <PrimaryButton
            onClick={handleSaveTask}
            styles={{ root: { marginRight: 8 } }}
          >
            Save
          </PrimaryButton>
          <DefaultButton onClick={handleClosePanel}>Cancel</DefaultButton>
        </div>
      </Panel>
    </div>
  );
};

export default SearchPanel;
