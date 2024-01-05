import { useState, useEffect,useContext } from "react";
import Pagination from "./Pagination";
import { TeamsFxContext } from "@microsoft/teamsfx-react";

import CardComponent from "./Card";

const Completed = (props) => {
  const [pages, setPages] = useState(1);
  const [selectedData, setselectedData] = useState([]);
  const [numberOfTask,setNumberOfTask]=useState(5)
  // const {
  //   teamsUserCredential,
  //   listTimeArry,
  //   loginuser,
  //   siteId,
  //   listToDoId,
  //   listToTaskEntryId,
  // } = useContext(TeamsFxContext);
  useEffect(() => {
    setselectedData([]);
    props?.listData.forEach((element) => {
      if (
        element?.fields.Status === "Completed"
        // (element.createdBy.user.email === loginuser.userPrincipalName ||
        //   element.fields.ReviewerMail === loginuser.userPrincipalName)
      ) {
        setselectedData((prev) => [...prev, element]);
      }
    });
  }, [props?.listData]);

  return (
    <div>
      {selectedData?.slice(pages * numberOfTask - numberOfTask, pages * numberOfTask).map((element) => {

        return (
          <div key={element.fields.id}>
            <CardComponent element={element} tabName={"Completed"} setCallReload={props?.setCallReload}/>
          </div>
        );
      })}
      <Pagination selectedData={selectedData} pages={pages}  setPages={setPages} numberOfTask={numberOfTask} setNumberOfTask={setNumberOfTask} selectedTab="Completed" />
    </div>
  );
};
export default Completed;
