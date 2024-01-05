import CardComponent from "./Card";
import { useState, useEffect,useContext } from "react";
import { Spinner } from "@fluentui/react-components";
import Pagination from "./Pagination";
import { TeamsFxContext } from "../Context";
import {getLogData} from "../util";

const OnGoing = (props) => {
  const [pages, setPages] = useState(1);
  const [numberOfTask,setNumberOfTask]=useState(5)
  const [onComlete,setOnComplete]=useState(false)
  const[onPause,setOnPause]=useState(false)
  const [selectedData, setselectedData] = useState([]);
  const[listTimeArry,setListTimeArry]=useState([])
  //////Try New api calls////////
  const {teamsUserCredential,siteId,listToDoId,listToTaskEntryId,}=useContext(TeamsFxContext)

useEffect(() => {
 
  if(teamsUserCredential&&siteId&&listToDoId&&listToTaskEntryId){
    getListData();
  }

}, [siteId,listToDoId,listToTaskEntryId])


const getListData=async()=>{
  const obj = {
    siteId: siteId,
    listid1: listToDoId,
    listid2: listToTaskEntryId,
  };
  const listData=await getLogData(teamsUserCredential,obj)
  setListTimeArry([])
  setListTimeArry(listData.listArray.value);
}
if(onPause){
  if(teamsUserCredential&&siteId&&listToDoId&&listToTaskEntryId){
    getListData();
   setOnPause(false)
  }
}

  /////////////
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
          listTimeArry={listTimeArry}
          setOnPause={setOnPause}
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
