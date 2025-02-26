import * as React from 'react';
import '../assets/Admin.css';
import Table from './Table';
import { Web } from 'sp-pnp-js';
import moment from 'moment';
import Expensetable from './Expensetable';
// import Salestable from './Salestable';
import { BarChart, Bar, Rectangle, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
// import { BarRectangleItem } from 'recharts/types/cartesian/Bar';

interface ChartData {
  month: string;
  expenses: number;
  sales?: number; 
}
const Admin: React.FC = () => {
  const [selectedMonth, setSelectedMonth] = React.useState(moment());
  const [userData, setUserData] = React.useState<any[]>([]);
  const [discontinuedUsers, setDiscontinuedUsers] = React.useState<any[]>([]);
  const [upcomingPayments, setUpcomingPayments] = React.useState<any[]>([]);
  const [dueDates, setDueDates] = React.useState<any[]>([]); 
  const [isTable, setIsTable] = React.useState(false);
 
  const [expensetableData, setExpensetableData] = React.useState<any[]>([]);
  const [currentData, setCurrentData] = React.useState<any[]>([]);
 const  [showExpenseTable, setShowExpenseTable] = React.useState(false);
 
  const [chartData, setChartData] = React.useState<ChartData[]>([]);
  const [totalExpenses, setTotalExpenses] = React.useState(0);
 
   const [upcomingAmountTotal, setUpcomingAmountTotal] = React.useState(0);
   const [PaymentdueTotal, setPaymentdueTotal] = React.useState(0);
  const [totalregistrations,settotalregistrations] =  React.useState<any[]>([]);
  const [totaldiscontinue, settotaldiscontinue] = React.useState<any[]>([]);
  const [totalearning,settotalearning] = React.useState(0);
  const [pendingPayments, setPendingPayments] = React.useState<any[]>([]);
   const [totalPending, setTotalPending] = React.useState<number>(0);
   const [totalTillExpense, setTotalTillExpense] = React.useState<number>(0);
 const [expenseData, setExpenseData] = React.useState<any[]>([]);
 const [showtotalExpenseTable, setShowtotalExpenseTable] = React.useState(false);
 const [chartSelectedMonth, setChartSelectedMonth] = React.useState<string | null>(null);
 const [showFilteredchartTable, setShowFilteredchartTable] = React.useState(false);
 const [expenseResData, setExpenseResData] = React.useState<any[]>([]);
 const [salesResData, setsalesResData] = React.useState<any[]>([]);
 const [expensetablechartData, setExpensetablechartData] = React.useState<any[]>([]);
 const [isTableVisible, setIsTableVisible] = React.useState(false);
 const [selectedsaleschartData, setSelectedsaleschartData] = React.useState<any[]>([]);
 const [chartSelectedsalesMonth, setChartSelectedsalesMonth] = React.useState<string>("");
const [amountReceived,setamountReceived] = React.useState(0);
const [clickedTiles, setClickedTiles] = React.useState(null);




  const fetchUserData = async (month: moment.Moment,isChartMonth = false) => {
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
          "DueDate",
          "BillAmount",
          "MedicalDetails",
          "BMIDate",
          "NextBMIDueDate",
          "MembershipNo",
          "MembershipPlan",
          "isActive",
          "PaymentDue",
          "PaymentHistory",
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
  
      res?.map((val: any) => {
        val.JoiningDate = moment(val.JoiningDate).format("DD/MM/YYYY");
        val.EndDate = val.EndDate ? moment(val.EndDate).format("DD/MM/YYYY") : "";
        val.BMIDate = val.BMIDate ? moment(val.BMIDate).format("DD/MM/YYYY") : "";
        val.NextBMIDueDate = val.NextBMIDueDate
          ? moment(val.NextBMIDueDate).format("DD/MM/YYYY")
          : "";
        val.DueDate = val.DueDate ? moment(val.DueDate).format("DD/MM/YYYY") : "";
      });

      setsalesResData(res)
  
      const startOfMonth = month.clone().startOf("month").startOf("day");
      const endOfMonth = month.clone().endOf("month").endOf("day");
      
      const filteredData = res.filter((item: any) => {
        if (!item?.JoiningDate) {
          return false;
        }
        const joiningMonth = moment(item.JoiningDate, "DD/MM/YYYY");
        return joiningMonth.isBetween(startOfMonth, endOfMonth, "day", "[]");
      });
  
      if (!isChartMonth) {
        setUserData(filteredData);
      }
     
      setCurrentData(filteredData);

      
      const discontinued: any[] = [];
      
      const filteredDiscontinue = res.filter((item: any) => {
        if (!item?.EndDate) {
          return false; 
        }
      
        const discontinuemonth = moment(item.EndDate, "DD/MM/YYYY");
        const isWithinMonth = discontinuemonth.isBetween(startOfMonth, endOfMonth, "day", "[]");
        const isOver = moment().isAfter(discontinuemonth, "day");
        
       
        const isInactive = item.isActive === false;
      
       
        if (isWithinMonth && isOver && isInactive) {
          discontinued.push(item); 
        }
      
        return isWithinMonth && isOver && isInactive;
      });

      const totalRegistrationsitem = res.filter((item:any) => item?.JoiningDate);

      let totalEarnings = 0;
      res.forEach((item: any) => {
        if (item?.AmountReceived) {
          totalEarnings += Number(item.AmountReceived);
        }
      });

      let EarningPermonth = 0;
      res.forEach((item: any) => {
        if (item?.AmountReceived && item?.JoiningDate) {
          const joiningDate = moment(item.JoiningDate, "DD/MM/YYYY");
          if (joiningDate.isBetween(startOfMonth, endOfMonth, null, '[]')) {
            EarningPermonth += Number(item.AmountReceived);
          }
        }
      });


      let pending = 0;
      const filteredPendingPayments = res.filter((item: any) => {
        if (item?.PaymentDue && Number(item.PaymentDue) > 0) {
          pending += Number(item.PaymentDue);
          return true;
        }
        return false;
      });
      
      
      
      const totalDiscontinueditem = res.filter((item: any) => {
        if (!item?.EndDate) return false;
      
        const discontinueDate = moment(item.EndDate, "DD/MM/YYYY");
        const isOver = moment().isAfter(discontinueDate, "day");
      
       
        const isInactive = item.isActive === false;
      
        return isOver && isInactive; 
      });
  
     
         
        let  totaPaymentdue = 0;
        const filteredPaymentdue = res.filter((item: any) => {
          if (!item?.DueDate) {
            return false; 
          }
        
          const PaymentdueMonth = moment(item.DueDate, "DD/MM/YYYY");
          const isWithinMonth = PaymentdueMonth.isBetween(startOfMonth, endOfMonth, "day", "[]");
          const isOverdue = moment().isAfter(PaymentdueMonth, "day"); 
          if (isWithinMonth && isOverdue) {
            const billAmount = Number(item?.BillAmount ?? 0);
            const paymentDue = Number(item?.PaymentDue ?? 0);
            totaPaymentdue += billAmount + paymentDue;
          }
        
          return isWithinMonth && isOverdue; 
        });
        
       
        
        
     
      let calculatedTotal = 0;
       const filteredByDueDate = res.filter((item: any) => {
        if (!item?.DueDate) {
          return false;
        }
        const dueMonth = moment(item.DueDate, "DD/MM/YYYY").subtract(7, "days");
        const isWithinMonth = dueMonth.isBetween(startOfMonth, endOfMonth, "day", "[]");
      
        if (isWithinMonth) {
          
          const billAmount = Number(item?.BillAmount ?? 0);
          const paymentDue = Number(item?.PaymentDue ?? 0);
          calculatedTotal += billAmount + paymentDue;
        }
      
        return isWithinMonth;
      });

      
      if (!isChartMonth) {

        setamountReceived(EarningPermonth)
        setPendingPayments(filteredPendingPayments);
      setTotalPending(pending);
      settotalearning(totalEarnings)
      settotalregistrations(totalRegistrationsitem);
      settotaldiscontinue(totalDiscontinueditem);
      setUpcomingAmountTotal(calculatedTotal);
      setPaymentdueTotal(totaPaymentdue)
      setDiscontinuedUsers(filteredDiscontinue);
      setUpcomingPayments(filteredByDueDate );
      setCurrentData(filteredData);
      setDueDates(filteredPaymentdue); 
      }
    
      

    } catch (error) {
      console.error("Error fetching user data:", error);
    }
  };

  const handlesalesBarClick = (data: any) => {
    const clickedMonth = moment(data.month, "MM/YY").format("MM/YYYY");
    const filteredsalesForMonth = salesResData.filter((item: any) => {
      const itemDate = moment(item.JoiningDate, "DD/MM/YYYY").format("MM/YYYY");
      return itemDate === clickedMonth ;
    });
    // const totalRegistrations = filteredsalesForMonth.length;
  
    setChartSelectedsalesMonth(clickedMonth);
    setSelectedsaleschartData(filteredsalesForMonth); 
    setShowFilteredchartTable(false); 
    setIsTableVisible(true);
     setIsTable(false)
 
   setShowExpenseTable(false);
  setShowtotalExpenseTable(false);

 
  };
  
  
  React.useEffect(() => {
    if (chartSelectedsalesMonth) {
   
      fetchUserData(moment(chartSelectedsalesMonth, "YYYY-MM"), true);
    } else {
     
      fetchUserData(selectedMonth);
    }
  }, [selectedMonth, chartSelectedsalesMonth]);
  


 const fetchExpenseAPIData = async (month: moment.Moment,isChartMonth = false) => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");
      const expenseres = await web.lists
        .getByTitle("Expenses")
        .items.select("Id","Amount", "ExpenseDate", "AmountType","SalesType","ExpenseType","PaymentMode","Comment","Modified",
         "Created",
          "Author/Id",
         "Author/Title",
         "Editor/Id",
         "Editor/Title").expand("Author,Editor")
        .top(4999)
        .get();

       const clientsRes = await web.lists
       .getByTitle("Clients")
       .items.select("AmountReceived", "JoiningDate")
       .top(4999)
       .get();
       setExpenseResData(expenseres); 

     
      const chartData: ChartData[] = [];
      // let totalExpenses = 0;
     

      const currentMonth = moment();
     const startMonth = currentMonth.clone().subtract(11, "months"); 

     for (let i = 0; i < 12; i++) {
      const monthYearLabel = startMonth.clone().add(i, "months").format("MM/YY");
      chartData.push({
        month: monthYearLabel,
        expenses: 0,
        sales: 0,
      });
    }

   

     const startOfMonth = month.clone().startOf("month").startOf("day").format("YYYY-MM-DD");
      const endOfMonth = month.clone().endOf("month").endOf("day").format("YYYY-MM-DD");

      

     
      expenseres.forEach((item: any) => {
        if (!item?.ExpenseDate) return;
      
        const itemDate = moment(item.ExpenseDate, "YYYY-MM-DD");
        if (!itemDate.isValid()) return;
      
        
        if (itemDate.isBetween(startMonth, currentMonth, undefined, "[]")) {
          const itemMonthName = itemDate.format("MM/YY");
      
          chartData.forEach((data) => {
            if (data.month === itemMonthName) {
              if (item.AmountType === "Expense") {
                data.expenses += item.Amount;
              }
              
            }
          });
        }
      
        
      
        
        
 
      });
      
      
      
      let TotalTillExpense = 0;
       const filteredTotalExpenses = expenseres.filter((item: any) => {
      if (item?.ExpenseDate && item.AmountType === "Expense") {
      const itemDate = moment(item.ExpenseDate, "YYYY-MM-DD");
       if (itemDate.isValid()) {
      TotalTillExpense += item.Amount;  
      return true;
    }
  }
  return false;
});
       
      
  
      
      clientsRes.forEach((client: any) => {
        if (!client?.JoiningDate) return;

      const joiningDate = moment(client.JoiningDate, "YYYY-MM-DD");

      if (joiningDate.isValid() && joiningDate.isBetween(startMonth, currentMonth, undefined, "[]")) {
        const itemMonthName = joiningDate.format("MM/YY");
    
        chartData.forEach((data) => {
          if (data.month === itemMonthName) {
            data.sales += client.AmountReceived || 0;  
          }
        });
      }
  
      });
  
  
      let totalExpenses = 0;
      const filteredExpenses = expenseres.filter((item: any) => {
        const itemDate = moment(item.ExpenseDate, "YYYY-MM-DD");
        if (
          item.AmountType === "Expense" &&
          itemDate.isValid() &&
          itemDate.isSameOrAfter(startOfMonth) &&
          itemDate.isSameOrBefore(endOfMonth)
        ) {
          totalExpenses += item.Amount || 0;
          return true;
        }
        return false;
      });

      if (!isChartMonth) {
        setTotalExpenses(totalExpenses);
      }
  
       setExpensetableData(filteredExpenses);
        setExpenseData(filteredTotalExpenses);
        setTotalTillExpense(TotalTillExpense );
      //  setTotalExpenses(totalExpenses); 
      setExpensetableData(filteredExpenses);
        
      setChartData(chartData);         
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  React.useEffect(() => {
    if (chartSelectedMonth) {
      
      fetchExpenseAPIData(moment(chartSelectedMonth, "YYYY-MM"), true);
    } else {
     
      fetchExpenseAPIData(selectedMonth);
    }
  }, [selectedMonth, chartSelectedMonth]);
  
  

 const handlePrevMonth = () => {
    const newMonth = selectedMonth.clone().subtract(1, 'month');
    setSelectedMonth(newMonth);
  };
  const handleNextMonth = () => {
    const newMonth = selectedMonth.clone().add(1, 'month');
    setSelectedMonth(newMonth);
  };

  const showAllUsers = () => {
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false)
    // setShowSalesTable(false)
   setIsTable(true);
    setCurrentData(userData);
  };
  const showtotalUsers = () => {
    setShowtotalExpenseTable(false);
    setShowFilteredchartTable(false);
    setShowExpenseTable(false);
  
    if (totalregistrations.length > 0) {
      
      const styledData = totalregistrations.map((user) => {
        const isInactive = user.isActive === false;
        const hasEnded = user.EndDate && moment(user.EndDate, "DD/MM/YYYY").isBefore(moment(), "day"); 
  
       
        return {
          ...user,
          highlightColor: isInactive && hasEnded ? "grey" : null,
        };
      });
  
      setIsTable(true); 
      setCurrentData(styledData); 
    }
  };
  
  const showPendingUsers = () => {
    // setShowSalesTable(false);
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false);
    if (pendingPayments.length > 0) {
      setIsTable(true);
      setCurrentData(pendingPayments);  
    }
  };
  const showtotalDiscontinuedUsers = () => {
    
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false)
    if (discontinuedUsers.length > 0) {
      setIsTable(true);
      setCurrentData(totaldiscontinue);
    }};

const showDiscontinuedUsers = () => {
    
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false)
    if (discontinuedUsers.length > 0) {
      setIsTable(true);
      setCurrentData(discontinuedUsers);
    }};

const showUpcomingPayments = () => {
   
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false)
    if (upcomingPayments.length > 0) {
      setIsTable(true);
      setCurrentData(upcomingPayments);
    }};

   const showDueDates = () => {
    
    setShowtotalExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowExpenseTable(false)
    if (dueDates.length > 0) {
      setIsTable(true);
      setCurrentData(dueDates);
    }};

    const handleShowExpenseTable = () => {
      setShowtotalExpenseTable(false)
      setShowFilteredchartTable(false);
    // setShowSalesTable(false)
    setIsTable(false);
    setShowExpenseTable(!showExpenseTable);
  };

  const handleShowTotalExpenseTable = () => {
    setShowExpenseTable(false)
    setShowFilteredchartTable(false);
    setShowtotalExpenseTable(!showtotalExpenseTable);
    setIsTable(false);
    setCurrentData(expenseData);  // Update table data with filtered results
  };

  const handleexpenseBarClick = (data: any) => {
    const clickedMonth = data.month;
    const filteredExpensesForMonth = expenseResData.filter((item: any) => {
      const itemDate = moment(item.ExpenseDate, "YYYY-MM-DD").format("MM/YY");
      return itemDate === clickedMonth && item.AmountType === "Expense";
    });
  
    setChartSelectedMonth(clickedMonth);
    setExpensetablechartData(filteredExpensesForMonth);  
    setShowFilteredchartTable(true);
    setIsTable(false); 
    setIsTableVisible(false);
    setShowExpenseTable(false); 
   setShowtotalExpenseTable(false); 
 

  };
  
  const CustomTooltip = ({ active, payload, label, }:any) => {
    if (active && payload && payload.length) {
     
      const clickedMonth = moment(label, "MM/YY").format("MM/YYYY");
      const totalRegistrations = salesResData.filter((item: any) => {
        const itemDate = moment(item.JoiningDate, "DD/MM/YYYY").format("MM/YYYY");
        return itemDate === clickedMonth;
      }).length;
      return (
        <div className="custom-tooltip" style={{ background: "#fff", padding: "10px", border: "1px solid #ccc", borderRadius: "5px" }}>
          <p style={{ margin: 0 }}>Month: {label}</p>
          <p style={{ margin: 0 }}>Sales: {payload[0].value}</p>
          <p style={{ margin: 0 }}> Registrations: {totalRegistrations}</p>
        </div>
      );
    }
    return null
  }
 

  const handleTileClick = (key:any, callback:any) => {
    setClickedTiles(key); // Update the active tile
    callback(); // Execute the corresponding function
  };
  return (
    <div className="admin-dashboard">
      <div className="text-end w-100 flex">
      <span
      style={{
      fontSize: '18px',
      fontWeight: '600',
      color: '#FFFFFF',
      margin: '0 16px',
    }}
  >
    {selectedMonth.format("MMM YYYY")} 
  </span>
  <button
    onClick={handlePrevMonth}
    style={{
      backgroundColor: '#87ceeb',
      border: 'none',
      padding: '8px 16px',
      borderRadius: '4px',
      fontWeight: 'bold',
      cursor: 'pointer',
      transition: 'background-color 0.3s ease',
    }}
  >
    Prev
  </button>
 <button
    onClick={handleNextMonth}
    style={{
      backgroundColor: '#87ceeb',
      border: 'none',
      padding: '8px 16px',
      borderRadius: '4px',
      fontWeight: 'bold',
      cursor: 'pointer',
      transition: 'background-color 0.3s ease',
      marginLeft: '8px',
    }}
  >
    Next
  </button>
       </div>

      <h1 className="admin-title">Admin Dashboard</h1>
      <div className="stats-container">
      <div
        className="stat-box"
        onClick={() => handleTileClick('totalRegistrations', showtotalUsers)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles === 'totalRegistrations' ? 'grey' : 'white',
        }}
      >
        Total Registrations<br />
        {totalregistrations.length}
        <span style={{ fontWeight: 'bold' }}> / </span>
        {'\u20B9'} {totalearning}
      </div>
      <div
        className="stat-box"
        onClick={() => handleTileClick('totalExpenses', handleShowTotalExpenseTable)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles ==='totalExpenses' ? 'grey' : 'white',
        }}
      >
        Total Expenses<br />
        {'\u20B9'}{totalTillExpense}
      </div>

      <div
        className="stat-box"
        onClick={() => handleTileClick('totalDiscontinued', showtotalDiscontinuedUsers)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles === 'totalDiscontinued' ? 'grey' : 'white',
        }}
      >
        Total Discontinued<br />
        {totaldiscontinue?.length || 0}
      </div>
       
      <div
        className="stat-box"
        onClick={() => handleTileClick('newRegistrations', showAllUsers)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles==='newRegistrations' ? 'grey' : 'white',
        }}
      >
        New Registrations<br />
        {userData.length} <span style={{ fontWeight: 'bold' }}> / </span> {'\u20B9'} {amountReceived}
      </div>
       
      <div
        className="stat-box"
        onClick={() => handleTileClick('pendingPayments', showPendingUsers)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles ==='pendingPayments' ? 'grey' : 'white',
        }}
      >
        Total Pending Payments<br />
        {'\u20B9'}{totalPending}
      </div>
      <div
        className="stat-box"
        onClick={() => handleTileClick('expenseAmount', handleShowExpenseTable)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles ==='expenseAmount' ? 'grey' : 'white',
        }}
      >
        Expenses Amount<br /> {'\u20B9'} {totalExpenses}
      </div>
      <div
        className="stat-box"
        onClick={() => handleTileClick('discontinued', showDiscontinuedUsers)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles==='discontinued' ? 'grey' : 'white',
        }}
      >
        Discontinued<br />
        {discontinuedUsers.length}
      </div>
      <div
        className="stat-box"
        onClick={() => handleTileClick('upcomingPayments', showUpcomingPayments)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles==='upcomingPayments' ? 'grey' : 'white',
        }}
      >
        Upcoming Payments<br />
        {'\u20B9'}{upcomingAmountTotal}
      </div>

      <div
        className="stat-box"
        onClick={() => handleTileClick('paymentDues', showDueDates)}
        style={{
          cursor: 'pointer',
          backgroundColor: clickedTiles==='paymentDues'? 'grey' : 'white',
        }}
      >
        Payment Dues<br />
        {'\u20B9'}{PaymentdueTotal}
      </div>

      </div>
      <div className="charts-container">
       {/* Sales Graph */}
  <div className="chart-box">
    <ResponsiveContainer width="100%" height={300}>
      <BarChart
        data={chartData}
        margin={{
          top: 5,
          right: 30,
          left: 20,
          bottom: 5,
        }}
      
      >
         <CartesianGrid strokeDasharray="3 3" />
        <XAxis
          dataKey="month"
          tickFormatter={(month, index) => (index % 2 === 0 ? month : "")}
        />
        <YAxis />
        <Tooltip content={<CustomTooltip salesResData={salesResData} />} />


        <Legend />
        <Bar
          dataKey="sales"
          fill="#800080"
          activeBar={<Rectangle fill="pink" stroke="blue" />}
          name="Sales"
          onClick={(data) => handlesalesBarClick(data)}
        />
      </BarChart>
    </ResponsiveContainer>
  </div>
{/* Expenses Graph */}
  <div className="chart-box">
    <ResponsiveContainer width="100%" height={300}>
      <BarChart
        data={chartData}
        margin={{
          top: 5,
          right: 30,
          left: 20,
          bottom: 5,
        }}
      >
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis
          dataKey="month"
          tickFormatter={(month, index) => (index % 2 === 0 ? month : "")}
        />
        <YAxis />
        <Tooltip />
        <Legend />
        <Bar
          dataKey="expenses"
          fill="#00C49F"
          activeBar={<Rectangle fill="gold" stroke="purple" />}
          name="Expenses"
          onClick={(data) => handleexpenseBarClick(data)}
        />
      </BarChart>
    </ResponsiveContainer>
  </div>
</div>
{isTable && (
        <div className="table-container">
          <Table tableData={currentData} fetchUserData={fetchUserData} selectedMonth={selectedMonth}  />
        </div>                                                           
      )}
       {isTableVisible && (
        <div className="table-container">
          <Table tableData={Array.isArray(selectedsaleschartData) ? selectedsaleschartData:[] } 
          fetchUserData={fetchUserData} 
          chartSelectedsalesMonth={chartSelectedsalesMonth}
          selectedMonth={selectedMonth} 
     />
        </div>
      )} 

{showFilteredchartTable && (
  <div className="expense-table-container">
    <Expensetable 
      expenseData={Array.isArray(expensetablechartData) ? expensetablechartData : []} 
      fetchExpenseAPIData={fetchExpenseAPIData} 
      chartselectedMonth={chartSelectedMonth} 
       selectedMonth={selectedMonth} 
    />
  </div>
)} 
{ showExpenseTable &&(
  <div className="expense-table-container">
    <Expensetable 
            expenseData={Array.isArray(expensetableData) ? expensetableData : []}
            fetchExpenseAPIData={fetchExpenseAPIData}
            selectedMonth={selectedMonth}   
            chartselectedMonth={chartSelectedMonth}   />
  </div>
)}

 {showtotalExpenseTable && (
  <div className="total-expense-table-container">
    <Expensetable
      expenseData={expenseData}
      fetchExpenseAPIData={fetchExpenseAPIData}
      selectedMonth={selectedMonth}
      chartselectedMonth={chartSelectedMonth}
    />
  </div>
)}
    </div>
    
  );
};

export default Admin;

// const Table: React.FC<TableProps> = ({ tableData, fetchUserData, selectedMonth, chartSelectedsalesMonth }) =>