import CardComponent from "./Card";
import { useState, useEffect } from "react";
import { Spinner } from "@fluentui/react-components";
import Pagination from "./Pagination";

const OnGoing = (props) => {
  const [pages, setPages] = useState(1);
  const [numberOfTask,setNumberOfTask]=useState(5)
  const [onComlete,setOnComplete]=useState(false)
  const [selectedData, setselectedData] = useState([]);
  useEffect(() => {
    setselectedData([]);
    props?.listData?.forEach((element) => {
      if (
        new Date(element.fields?.StartDate) <= new Date() &&
        element.fields.Status !== "Completed"
      ) {
        setselectedData((prev) => [...prev, element]);
      }
    });
  }, [props?.listData]);
  return (<div>
    {onComlete?(<div style={{width:"100%",height:"100%"}}><Spinner label="Completing The Task" labelPosition="below" /> </div>): 
    <div>
    {selectedData?.slice(pages * numberOfTask - numberOfTask, pages * numberOfTask).map((element) => {
      return (
        <div  key={element.fields.id} >
          <CardComponent
          
          setOnComplete={setOnComplete}
            element={element}
            setCallReload={props.setCallReload}
            tabName={"OnGoing"}
          />
        </div>
      );
    })}
    
  </div>}
   <Pagination selectedData={selectedData} pages={pages}  setPages={setPages} numberOfTask={numberOfTask} setNumberOfTask={setNumberOfTask} selectedTab="OnGoing" />
   
  </div>
  );
};
export default OnGoing;
