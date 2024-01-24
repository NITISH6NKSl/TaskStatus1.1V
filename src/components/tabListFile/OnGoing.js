import CardComponent from "./Card";
import { useState, useEffect,useContext } from "react";
import { Spinner } from "@fluentui/react-components";
import Pagination from "./Pagination";
import { TeamsFxContext } from "../Context";
import {getLogData,getListDataCall} from "../util";

const OnGoing = (props) => {
  const [pages, setPages] = useState(1);
  const [numberOfTask,setNumberOfTask]=useState(5)
  const [onComlete,setOnComplete]=useState(false)
  const[onPause,setOnPause]=useState(false);
  const [taskLoad,setTaskLoad]=useState(false)
  const [ongoingPageLoad,setOngoingPageLoad]=useState(true)
  const [selectedData, setselectedData] = useState([]);
  const[listTimeArry,setListTimeArry]=useState([])
  
  //////Try New api calls////////
  const {teamsUserCredential,siteId,listToDoId,listToTaskEntryId,}=useContext(TeamsFxContext)

useEffect(() => {
 
  if(teamsUserCredential&&siteId&&listToDoId&&listToTaskEntryId){
    getListData();
  }

}, [siteId,listToDoId,listToTaskEntryId])
useEffect(() => {
  if(teamsUserCredential&&siteId&&listToDoId&&listToTaskEntryId){
    getTaskData();
  }
}, [siteId,listToDoId,])
// useEffect(()=>{
//  SetlistDataOnSearch()    
// },[])
// const SetlistDataOnSearch=()=>{
//   if(props?.SearchData.length>0){
//     settaskListData(props?.list)
//   }
// }
const getTaskData= async()=>{
  const obj = {
    siteId: siteId,
    listid1: listToDoId,
  };
  const taskdata=await getListDataCall(teamsUserCredential,obj)
  setselectedData([]);
  taskdata.graphClientMessage.value?.forEach((element) => {
    if (
      new Date(element.fields?.StartDate) <= new Date() &&
      element.fields.Status !== "Completed"
    ) {
      setselectedData((prev) => [...prev, element]);
    }
  });
  setTaskLoad(false)
  setOngoingPageLoad(false)
  // settaskListData()
}
console.log("This is a data of task list data in use State",props.SearchData)
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
const callLogData=async(check)=>{
  if(check){
    setOnPause(true)
    if(teamsUserCredential&&siteId&&listToDoId&&listToTaskEntryId){
      await getListData();
      setOnPause(false)
      return "Success"
    }
    else{
      return "Failed"
    }
  }

}
if(taskLoad &&teamsUserCredential&&siteId&&listToDoId){
  getTaskData() 
}
  /////////////
  return (<div>
    {ongoingPageLoad?<><Spinner label="Data loading" labelPosition="below" style={{display:"flex",justifyContent:"center",height:"65%",alignItems:"center"}}></Spinner></>:<>
    {onComlete?(<div style={{width:"100%",height:"100%"}}><Spinner label="Completing the task" labelPosition="below" /> </div>): 
    <div>
    {(props.SearchData.length>0?(props?.SearchData):selectedData)?.slice(pages * numberOfTask - numberOfTask, pages * numberOfTask).map((element) => {
        if (
          new Date(element.fields?.StartDate) <= new Date() &&
          element.fields.Status !== "Completed"
        ) {
      return (
        <div  key={element.fields.id} >
          <CardComponent
          setTaskLoad={setTaskLoad}
          // setDialogVisibility={props?.setDialogVisibility}
          listTimeArry={listTimeArry}
          callLogData={callLogData}
          setOnComplete={setOnComplete}
          element={element}
          setCallReload={props.setCallReload}
          tabName={"OnGoing"}
          />
        </div>
          );
      }
    })}
  </div>}
   <Pagination selectedData={(props.SearchData.length>0?(props?.SearchData):selectedData)} pages={pages}  setPages={setPages} numberOfTask={numberOfTask} setNumberOfTask={setNumberOfTask} selectedTab="OnGoing" />
   </>
    }
  </div>
     
    
    
  );
};
export default OnGoing;
