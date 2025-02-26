import * as React from "react";
import { Web } from "sp-pnp-js";
import "../Trainer/Trainer_Dashboard.css";
import { RxCross2 } from "react-icons/rx";
// import { MdEditCalendar } from "react-icons/md";
import SearchPanel from "./SearchPanel";

interface DietChart {
  mealTime: string;
  mealDescription: string;
  caloriesPortion: string;
  timeSlot: string;
  dayOfWeek: string;
}

interface BMIDetails {
  date: string;
  name: string;
  contact: string;
  height: number | undefined;
  weight: number | undefined;
  age: number | undefined;
  vFat: number | undefined; // Visceral Fat
  kcal: number | undefined; // Calorie Intake
  bodyAge: number | undefined;
  fat: number | undefined;
  bmi: number | undefined;
  subWholeBody: number | undefined;
  subTrunk: number | undefined;
  subArms: number | undefined;
  subLegs: number | undefined;
  sMusclesWholeBody: number | undefined;
  sMusclesTrunk: number | undefined;
  sMusclesArms: number | undefined;
  sMusclesLegs: number | undefined;
}

interface ProgressReport {
  date: "";
  type: "BMI";
  value: "";
  status: string;
}

interface Feedback {
  type: "Complaint" | "Suggestion";
  description: string;
}

const Trainer_Dashboard: React.FC = () => {
  const [data, setData] = React.useState({
    FullName: "",
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
    Feedback: [
      {
        type: "Complaint",
        description: "",
      },
    ],
  });
  const [currentDietEntry, setCurrentDietEntry] = React.useState<DietChart>({
    mealTime: "",
    mealDescription: "",
    caloriesPortion: "",
    timeSlot: "",
    dayOfWeek: "",
  });
  const [storeData, setStoreData] = React.useState([]);
  const [searchItem, setSearchItem] = React.useState("");
  const [searchResult, setSearchResult] = React.useState<any[]>([]);
  const [showModal, setShowModal] = React.useState<boolean>(false);
  const [selectedResults, setSelectedResults] = React.useState(null);
  // const [alertTriggered, setAlertTriggered] = React.useState(false);
  const [bmiDetails, setBmiDetails] = React.useState<BMIDetails>({
    date: "",
    name: "",
    contact: "",
    height: undefined,
    weight: undefined,
    age: undefined,
    vFat: undefined,
    kcal: undefined,
    bodyAge: undefined,
    fat: undefined,
    bmi: undefined,
    subWholeBody: undefined,
    subTrunk: undefined,
    subArms: undefined,
    subLegs: undefined,
    sMusclesWholeBody: undefined,
    sMusclesTrunk: undefined,
    sMusclesArms: undefined,
    sMusclesLegs: undefined,
  });
  const [feedback, setFeedback] = React.useState<Feedback>({
    type: "Complaint",
    description: "",
  });

  const [progressReport, setProgressReport] = React.useState<ProgressReport>({
    date: "",
    type: "BMI",
    value: "",
    status: "",
  });
  const [showTable, setShowTable] = React.useState(false);
  const [showBmiInSevDays, setShowBmiInSevDays] = React.useState(false);
  const [todayBMIDue, setTodayBMIDue] = React.useState([]);
  const [nextSevenDaysBMIDue, setNextSevenDaysBMIDue] = React.useState([]);

  //Handle Change BMI index
  const handleChange = (
    field: keyof BMIDetails,
    value: string | number | undefined
  ) => {
    setBmiDetails((prev) => ({
      ...prev,
      [field]: value,
    }));

    // Update the corresponding entry in `data.BMIDetails`
    setData((prevData) => ({
      ...prevData,
      BMIDetails: prevData.BMIDetails.map((item, idx) =>
        idx === 0 // Assuming you're updating the first item in BMIDetails
          ? {
              ...item,
              [field]: value,
            }
          : item
      ),
    }));
  };

  const handleFeedbackChange = (field: keyof Feedback, value: string) => {
    setFeedback((prev) => ({
      ...prev,
      [field]: value,
    }));
    setData((prevData) => ({
      ...prevData,
      Feedback: prevData.Feedback.map((item, idx) =>
        idx === 0 // Assuming you're updating the first item in BMIDetails
          ? {
              ...item,
              [field]: value,
            }
          : item
      ),
    }));
  };

  // const calculateBMI = (weight: number, height: number): number | null => {
  //   if (!weight || !height || height <= 0) return null;
  //   return parseFloat((weight / (height * height)).toFixed(2));
  // };

  // const handleChange = (
  //   field: keyof BMIDetails,
  //   value: string | number | undefined
  // ) => {
  //   setBmiDetails((prev) => ({
  //     ...prev,
  //     [field]: value,
  //   }));

  //   // Dynamically calculate BMI if both weight and height are provided
  //   let calculatedBMI: number | null = null;

  //   if (
  //     (field === "weight" || field === "height") &&
  //     bmiDetails.weight &&
  //     bmiDetails.height
  //   ) {
  //     const weight =
  //       field === "weight" ? Number(value) : Number(bmiDetails.weight);
  //     const height =
  //       field === "height" ? Number(value) : Number(bmiDetails.height);
  //     calculatedBMI = calculateBMI(weight, height);
  //   }

  //   setBmiDetails((prev) => ({
  //     ...prev,
  //     [field]: value,
  //   }));

  //   // Update the corresponding entry in `data.BMIDetails`
  //   setData((prevData) => ({
  //     ...prevData,
  //     BMIDetails: prevData.BMIDetails.map((item, idx) =>
  //       idx === 0 // Assuming you're updating the first item in BMIDetails
  //         ? {
  //             ...item,
  //             [field]: value,
  //             ...(field === "height" || field === "weight"
  //               ? { bmi: calculatedBMI ? calculatedBMI.toString() : "" } // Convert BMI to string if necessary
  //               : {}),
  //           }
  //         : item
  //     ),
  //   }));
  // };

  const calculateStatus = (type: string, value: string): string => {
    let status = "";
    const numericValue = parseFloat(value);

    if (type === "BMI" && !isNaN(numericValue)) {
      if (numericValue < 18.5) status = "Underweight";
      else if (numericValue >= 18.5 && numericValue <= 24.9) status = "Normal";
      else if (numericValue >= 25 && numericValue <= 29.9)
        status = "Overweight";
      else status = "Obesity";
    } else if (type === "Weight" && !isNaN(numericValue)) {
      status = numericValue > 70 ? "Needs Attention" : "Healthy";
    }

    return status;
  };

  //handle Change Progress Report
  const handleChangeP = (
    field: keyof ProgressReport,
    value: string | number | undefined
  ) => {
    // Calculate status dynamically for specific fields
    let updatedStatus: any = progressReport.status;
    // let BW:any = progressReport.type;

    if (progressReport.type === "BMI" || progressReport.type === "Weight") {
      updatedStatus = calculateStatus(progressReport.type, String(value));
    }

    setProgressReport((prev) => ({
      ...prev,
      [field]: value,
      status: updatedStatus,
    }));
    // Update the corresponding entry in `data.ProgressReport`
    setData((prevData) => ({
      ...prevData,
      ProgressReport: prevData.ProgressReport.map((item, idx) =>
        idx === 0 // Assuming you're updatin the first item in ProgressReport
          ? {
              ...item,
              [field]: value,
              status: updatedStatus,
            }
          : item
      ),
    }));
  };

  const handleDietEntryChange = (key: string, value: string) => {
    setCurrentDietEntry((prev) => ({
      ...prev,
      [key]: value,
    }));
    setData((prev) => ({
      ...prev,
      DietChart: prev.DietChart.map((item, idx) =>
        idx === 0 ? { ...item, [key]: value } : item
      ),
    }));
  };

  const addDietEntry = () => {
    // Validate that all fields are provided
    if (
      !currentDietEntry.mealTime ||
      !currentDietEntry.mealDescription ||
      !currentDietEntry.caloriesPortion ||
      !currentDietEntry.timeSlot ||
      !currentDietEntry.dayOfWeek
    ) {
      console.error("All fields in currentDietEntry must be filled.");
      return;
    }

    // Replace the DietChart with the new entry
    setData((prevData) => ({
      ...prevData,
      DietChart: [
        {
          mealTime: currentDietEntry.mealTime,
          mealDescription: currentDietEntry.mealDescription,
          caloriesPortion: currentDietEntry.caloriesPortion,
          timeSlot: currentDietEntry.timeSlot,
          dayOfWeek: currentDietEntry.dayOfWeek,
        },
      ],
    }));
  };

  const fetchAPIData = async () => {
    const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
    try {
      const res = await web.lists.getByTitle("Clients").items.get();
      setStoreData(res);
      console.log("Fetched Data", res);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  const handleSearch = (value: string) => {
    setSearchItem(value);
    if (value.trim() === "") {
      setSearchResult([]); // Clear dropdown if input is empty
      return;
    }
    // Filter the data (assuming `data` is your dataset)
    const filtered = storeData?.filter((item: any) => {
      if (item?.FullName !== null && item?.FullName !== undefined) {
        return item?.FullName.toLowerCase().includes(value.toLowerCase());
      }
    });
    setSearchResult(filtered);
  };
  console.log(storeData);
  console.log(selectedResults);
  console.log(setData);
  React.useEffect(() => {
    fetchAPIData();
  }, []);

  // Add new item to the list
  const addFunction = async () => {
    const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
    try {
      const response = await web.lists
        .getById("3A9C0B25-B14D-4277-99CB-D63FCFF5FD3F")
        .items.add({
          BMIDetails: JSON.stringify(data.BMIDetails),
          DietChart: JSON.stringify(data.DietChart),
          ProgressReport: JSON.stringify(data.ProgressReport),
          Feedback: JSON.stringify(data.Feedback),
        });
      console.log("Add Function Response", response);
      await fetchAPIData();
    } catch (error) {
      console.error("Error adding item:", error);
    }
    setData({
      FullName: "",
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
      Feedback: [
        {
          type: "Complaint",
          description: "",
        },
      ],
    });
  };

  const handleSubmit = async () => {
    try {
      // Wait for state to update fully by using a callback function in setState
      await new Promise<void>((resolve) => {
        setData((prevData) => {
          addDietEntry();
          resolve();
          return prevData;
        });
      });
      await addFunction();
    } catch (error) {
      console.error("Error in submitting data:", error);
    } finally {
      // Reset `currentDietEntry` after submission
      setCurrentDietEntry({
        mealTime: "",
        mealDescription: "",
        caloriesPortion: "",
        timeSlot: "",
        dayOfWeek: "",
      });
    }
  };
  console.log(showModal);

  const handleClick = (result: any) => {
    setSelectedResults(result);
    setSearchItem(result.FullName); // Set the input value to the selected result
    setSearchResult([]); // Clear the dropdown
    setShowModal(true);
  };
  const handleClose = () => {
    setShowTable(false);
  };

  React.useEffect(() => {
    // Filter items where the difference is 7 days
    const filterBMIDue = storeData.filter((item: any) => {
      if (item.BMIDate) {
        const bmiDate = new Date(item.BMIDate); // Convert BMIDate to Date object
        const currentDate = new Date(); // Current date
        const timeDiff = bmiDate.getTime() - currentDate.getTime(); // Difference in milliseconds
        const dayDiff = Math.ceil(timeDiff / (1000 * 60 * 60 * 24)); // Convert to days
        return dayDiff === 7; // Check if the difference is exactly 7 days
      }
      return false;
    });

    setTodayBMIDue(filterBMIDue); // Store the filtered items in state
  }, [storeData]);

  React.useEffect(() => {
    // Filter items where the NextBMIDueDate is within the next 7 days
    const filterNextSevenDaysBMIDue = storeData.filter((item: any) => {
      if (item.NextBMIDueDate) {
        const nextBmiDueDate = new Date(item.NextBMIDueDate); // Convert NextBMIDueDate to Date object
        const currentDate = new Date(); // Current date
        const timeDiff = nextBmiDueDate.getTime() - currentDate.getTime(); // Difference in milliseconds
        const dayDiff = Math.ceil(timeDiff / (1000 * 60 * 60 * 24)); // Convert to days
        return dayDiff >= 0 && dayDiff <= 7; // Check if the difference is between 0 and 7 days (inclusive)
      }
      return false;
    });

    setNextSevenDaysBMIDue(filterNextSevenDaysBMIDue); // Store the filtered items in state
  }, [storeData]);

  return (
    <div className="main">
      <h1 className="title">Trainer Dashboard</h1>
      <div className="stats-container">
        <div className="stat-box" onClick={() => setShowTable(true)}>
          <p className="stat-text">BMI Dues Today</p>
          <h2 className="stat-number">{todayBMIDue.length}</h2>
        </div>
        {showTable && (
          <div className="modal-overlay">
            <div className="modal-content">
              <button className="close-button" onClick={handleClose}>
                x
              </button>
              <h2 className="modal-title">BMI Dues Today</h2>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>First Name</th>
                      <th>Last Name</th>
                      <th>Full Name</th>
                      <th>Last BMI Date</th>
                      <th>Last BMI Details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {todayBMIDue?.map((item: any, index: number) => (
                      <tr key={index}>
                        <td>{item.FirstName}</td>
                        <td>{item.Title}</td>
                        <td>{item.FullName}</td>
                        <td>{item.BMIDate}</td>
                        <td>{item.BMIDetails?.[0]?.details || "N/A"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
        <div className="stat-box" onClick={() => setShowBmiInSevDays(true)}>
          <p className="stat-text">BMI Dues in Next 7 Days</p>
          <h2 className="stat-number">{nextSevenDaysBMIDue.length}</h2>
        </div>
        {showBmiInSevDays && (
          <div className="modal-overlay">
            <div className="modal-content">
              <button
                className="close-button"
                onClick={() => setShowBmiInSevDays(false)}
              >
                x
              </button>
              <h2 className="modal-title">BMI Dues in Next 7 Days</h2>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>First Name</th>
                      <th>Last Name</th>
                      <th>Full Name</th>
                      <th>Last BMI Date</th>
                      <th>Last BMI Details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {nextSevenDaysBMIDue?.map((item: any, index: number) => (
                      <tr key={index}>
                        <td>{item.FirstName}</td>
                        <td>{item.Title}</td>
                        <td>{item.FullName}</td>
                        <td>{item.BMIDate}</td>
                        <td>{item.BMIDetails?.[0]?.details || "N/A"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
        <div className="stat-box">
          <p className="stat-text">PTs Assigned Today</p>
          <h2 className="stat-number">2</h2>
        </div>
        <div className="stat-box">
          <p className="stat-text">PTs Assigned in Next 7 Days</p>
          <h2 className="stat-number">4</h2>
        </div>
      </div>
      <div className="form-container">
        <div className="form-group">
          <label htmlFor="feedbackType" className="label">
            FeedBack
          </label>
          <table className="feedback-table">
            <thead>
              <tr>
                <th>Type</th>
                <th>Description</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>
                  <select
                    className="select"
                    id="feedbackType"
                    value={feedback.type}
                    onChange={(e) =>
                      handleFeedbackChange(
                        "type",
                        e.target.value as "Complaint" | "Suggestion"
                      )
                    }
                  >
                    <option value="complaint">Complaint</option>
                    <option value="suggestion">Suggestion</option>
                  </select>
                </td>
                <td>
                  <textarea
                    id="feedbackDescription"
                    className="textarea"
                    placeholder="Describe your complaint or suggestion here"
                    rows={5}
                    cols={127}
                    onChange={(e) =>
                      handleFeedbackChange("description", e.target.value)
                    }
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        {/* BMI Details */}
        <div className="form-group">
          <label htmlFor="enterBMIDetails" className="label">
            Enter BMI Details
          </label>
          <table className="form-table">
            <thead>
              <tr>
                <th>Parameter</th>
                <th>value</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Date</td>
                <td>
                  <input
                    type="date"
                    className="input"
                    placeholder="DD-MM-YYYY"
                    id="date"
                    value={bmiDetails.date}
                    onChange={(e) => handleChange("date", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Name</td>
                <td>
                  <input
                    type="text"
                    className="input"
                    id="name"
                    placeholder="Enter Name"
                    value={bmiDetails.name}
                    onChange={(e) => handleChange("name", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Contact</td>
                <td>
                  <input
                    type="text"
                    className="input"
                    id="contact"
                    placeholder="Enter Contact"
                    value={bmiDetails.contact}
                    onChange={(e) => handleChange("contact", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Weight (kg)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="weight"
                    placeholder="Enter Weight"
                    value={bmiDetails.weight}
                    onChange={(e) => handleChange("weight", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Height (cm)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="height"
                    placeholder="Enter Height"
                    value={bmiDetails.height}
                    onChange={(e) => handleChange("height", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Age</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="age"
                    placeholder="Enter Age"
                    value={bmiDetails.age}
                    onChange={(e) => handleChange("age", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Visceral Fat (%)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="vFat"
                    placeholder="Enter Visceral Fat"
                    value={bmiDetails.vFat}
                    onChange={(e) => handleChange("vFat", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Kcal</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="kcal"
                    placeholder="Enter Kcal"
                    value={bmiDetails.kcal}
                    onChange={(e) => handleChange("kcal", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Body Age</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="bodyAge"
                    placeholder="Enter Body Age"
                    value={bmiDetails.bodyAge}
                    onChange={(e) => handleChange("bodyAge", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Fat (%)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="fat"
                    placeholder="Enter Fat Percentage"
                    value={bmiDetails.fat}
                    onChange={(e) => handleChange("fat", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>BMI</td>
                <td>
                  <input
                    type="text"
                    className="input"
                    id="bmi"
                    readOnly
                    value={bmiDetails.bmi}
                  />
                </td>
              </tr>
              <tr>
                <td>Sub Whole Body</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="subWholeBody"
                    placeholder="Enter Sub Whole Body"
                    value={bmiDetails.subWholeBody}
                    onChange={(e) =>
                      handleChange("subWholeBody", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Sub Trunk</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="subTrunk"
                    placeholder="Enter Sub Trunk"
                    value={bmiDetails.subTrunk}
                    onChange={(e) => handleChange("subTrunk", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Sub Arms</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="subArms"
                    placeholder="Enter Sub Arms"
                    value={bmiDetails.subArms}
                    onChange={(e) => handleChange("subArms", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Sub Legs</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="subLegs"
                    placeholder="Enter Sub Legs"
                    value={bmiDetails.subLegs}
                    onChange={(e) => handleChange("subLegs", e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <td>Skeletal Muscles (Whole Body)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="sMusclesWholeBody"
                    placeholder="Enter Skeletal Muscles Whole Body"
                    value={bmiDetails.sMusclesWholeBody}
                    onChange={(e) =>
                      handleChange("sMusclesWholeBody", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Skeletal Muscles (Trunk)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="sMusclesTrunk"
                    placeholder="Enter Skeletal Muscles Trunk"
                    value={bmiDetails.sMusclesTrunk}
                    onChange={(e) =>
                      handleChange("sMusclesTrunk", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Skeletal Muscles (Arms)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="sMusclesArms"
                    placeholder="Enter Skeletal Muscles Arms"
                    value={bmiDetails.sMusclesArms}
                    onChange={(e) =>
                      handleChange("sMusclesArms", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Skeletal Muscles (Legs)</td>
                <td>
                  <input
                    type="number"
                    className="input"
                    id="sMusclesLegs"
                    placeholder="Enter Skeletal Muscles Legs"
                    value={bmiDetails.sMusclesLegs}
                    onChange={(e) =>
                      handleChange("sMusclesLegs", e.target.value)
                    }
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        {/* END BMI Chart */}
        <div className="form-group">
          <label htmlFor="searchClient" className="label">
            Search Client
          </label>
          <div className="bind">
            <input
              id="searchClient"
              className="input"
              type="text"
              placeholder="Search by Name or Contact"
              value={searchItem}
              onChange={(e) => handleSearch(e.target.value)}
            />
            {searchResult.length > 0 && (
              <>
                <span className="cross">
                  <RxCross2
                    className="clear-icon"
                    onClick={() => setSearchItem("")} // Clear the input when clicked
                  />
                </span>

                <ul className="dropdown">
                  {searchResult.map((result, index) => (
                    <li
                      key={index}
                      className="dropdown-item"
                      onClick={() => handleClick(result)}
                    >
                      {result.FullName}
                    </li>
                  ))}
                </ul>
              </>
            )}
            {showModal && (
              <SearchPanel
                selectedResults={selectedResults}
                onClose={() => setShowModal(false)}
              />
            )}
          </div>
        </div>
        <div className="progress-report-container">
          <label htmlFor="enterProgressReport" className="label">
            Enter Progress Report
          </label>
          <table className="form-table">
            <thead>
              <tr>
                <th>Parameter</th>
                <th>Value</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Date</td>
                <td>
                  <input
                    type="date"
                    value={progressReport.date}
                    onChange={(e) => handleChangeP("date", e.target.value)}
                    className="input"
                  />
                </td>
              </tr>
              <tr>
                <td>Type</td>
                <td>
                  <select
                    value={progressReport.type}
                    onChange={(e) =>
                      handleChangeP("type", e.target.value as "BMI" | "Weight")
                    }
                    className="input"
                  >
                    <option value="BMI">BMI</option>
                    <option value="Weight">Weight</option>
                  </select>
                </td>
              </tr>
              <tr>
                <td>{progressReport.type} Value</td>
                <td>
                  <input
                    type="number"
                    value={progressReport.value}
                    onChange={(e) => handleChangeP("value", e.target.value)}
                    placeholder={`Enter ${progressReport.type}`}
                    className="input"
                  />
                </td>
              </tr>
              <tr>
                <td>Status</td>
                <td>
                  <input
                    type="text"
                    value={progressReport.status}
                    readOnly
                    className={`input ${
                      progressReport.status === "Normal"
                        ? "status-normal"
                        : progressReport.status === "Underweight"
                        ? "status-underweight"
                        : progressReport.status === "Overweight"
                        ? "status-overweight"
                        : progressReport.status === "Obesity"
                        ? "status-obesity"
                        : ""
                    }`}
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        {/* <div className="form-group">
          <label htmlFor="workSchedule" className="label">
            Enter Work Schedule
          </label>
          <textarea
            id="workSchedule"
            className="textarea"
            placeholder="Enter work schedule of a candidate"
          />
        </div> */}
        <div className="form-group">
          <h3>Enter Diet Chart</h3>
          <table className="form-table">
            <thead>
              <tr>
                <th>Field</th>
                <th>Input Type</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Meal Time</td>
                <td>
                  <select
                    className="input-select"
                    value={currentDietEntry.mealTime}
                    onChange={(e) =>
                      handleDietEntryChange("mealTime", e.target.value)
                    }
                  >
                    <option value="">Select...</option>
                    <option>Breakfast</option>
                    <option>Lunch</option>
                    <option>Dinner</option>
                  </select>
                </td>
              </tr>
              <tr>
                <td>Meal Description</td>
                <td>
                  <textarea
                    className="input-textarea"
                    placeholder="Enter details of the meal"
                    value={currentDietEntry.mealDescription}
                    onChange={(e) =>
                      handleDietEntryChange("mealDescription", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Calories/Portion</td>
                <td>
                  <input
                    type="text"
                    className="input-textbox"
                    placeholder="Enter calories or portion size"
                    value={currentDietEntry.caloriesPortion}
                    onChange={(e) =>
                      handleDietEntryChange("caloriesPortion", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Time Slot</td>
                <td>
                  <input
                    type="text"
                    className="input-textbox"
                    placeholder="Enter the time"
                    value={currentDietEntry.timeSlot}
                    onChange={(e) =>
                      handleDietEntryChange("timeSlot", e.target.value)
                    }
                  />
                </td>
              </tr>
              <tr>
                <td>Day of Week</td>
                <td>
                  <select
                    className="input-select"
                    value={currentDietEntry.dayOfWeek}
                    onChange={(e) =>
                      handleDietEntryChange("dayOfWeek", e.target.value)
                    }
                  >
                    <option value="">Select...</option>
                    <option>Monday</option>
                    <option>Tuesday</option>
                    <option>Wednesday</option>
                    <option>Thursday</option>
                    <option>Friday</option>
                    <option>Saturday</option>
                    <option>Sunday</option>
                  </select>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <div className="submit-button-container">
          <button className="submit-button" onClick={handleSubmit}>
            Submit
          </button>
        </div>
      </div>
    </div>
  );
};
export default Trainer_Dashboard;
