import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";
import CardComponent from "./Card";
import Pagination from "./Pagination";

const UpComing = (props) => {
  const [pages, setPages] = useState(1);
  const [selectedData, setselectedData] = useState([]);
  const [numberOfTask,setNumberOfTask]=useState(5)
  useEffect(() => {
    setselectedData([]);
    props?.listData.forEach((element) => {
      // console.log("This is created by", element.createdBy.email);
      // console.log("This is Reviwer in ", element.fields.Reviwer);
      if (
        element?.fields.Status !== "Completed" &&
        new Date(element.fields?.StartDate) > new Date()
        // (element.createdBy.user.email === loginuser.userPrincipalName ||
        //   element.fields.ReviewerMail === loginuser.userPrincipalName)
      ) {
        setselectedData((prev) => [...prev, element]);
      }
    });
  }, [props?.listData]);
  console.log("this is selected data ", selectedData);
  console.log("This is a data");
  return (
    <div>
      {selectedData?.slice(pages * numberOfTask - numberOfTask, pages * numberOfTask).map((element) => {
        return (
          <div key={element.fields.id}>
            <CardComponent
              element={element}
              style={{ justifyContent: "flex-end", paddingRight: "20px" }}
              tabName={"UpComing"}
            />
          </div>
        );
      })}
         <Pagination selectedData={selectedData} pages={pages}  setPages={setPages} numberOfTask={numberOfTask} setNumberOfTask={setNumberOfTask} selectedTab="UpComing" />
    </div>
  );
};
export default UpComing;
