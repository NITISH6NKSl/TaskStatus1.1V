import { Button,Dropdown,Option, Tooltip,Text } from "@fluentui/react-components";
import {ArrowNext24Regular,ArrowPrevious24Regular,IosArrowRtl24Regular,IosArrowLtr24Regular} from "@fluentui/react-icons";
import { useState,useEffect } from "react";

const NumberTaskArray=[1,3,5,7,9,11]
const Pagination =(props)=>{
    const [firstPage,setFirstPage]=useState()
    const [lastPage,setLastPage]=useState()
    useEffect(() => {
        setFirstPage(1)
        setLastPage((Math.ceil(props?.selectedData?.length / props?.numberOfTask)))
     
    }, [props])
const selectPagehandler = (e, selectedpage) => {
        e.preventDefault();
    
        if (
          selectedpage >= 1 &&
          selectedpage <= Math.ceil(props?.selectedData.length / props?.numberOfTask) &&
          selectedpage !== props?.pages
        ) {
          props?.setPages(selectedpage);
          
        }
      };
const handleTaskNumber=(taskNumber)=>{
            
            props?.setNumberOfTask(Number(taskNumber))
            props?.setPages(firstPage)
    }
    // const checkSelectedTab=()=>{
    //   switch(props?.selectedTab){
    //     case "OnGoing":
    //       return "On Going"
    //     case "UpComing":
    //       return "Up Coming"
    //     case "Completed":
    //       return "Completed"
    //     default:
    //       return ""
    //   }
    // }

///////
const totalPages = Math.ceil(props?.selectedData.length / props?.numberOfTask);

const renderPageNumbers = () => {
  const currentPage = props.pages;
  const pageNumbers = [];
  const maxPagesToShow = 5;

  if (totalPages <= maxPagesToShow) {
    for (let i = 1; i <= totalPages; i++) {
      pageNumbers.push(
        <span
          className={currentPage === i ? "selectedPage" : ""}
          onClick={(e) => selectPagehandler(e, i)}
          key={i}
        >
          {i}
        </span>
      );
    }
  } else {
    const middlePage = Math.floor(maxPagesToShow / 2) + 1;
    let startPage = currentPage - middlePage + 1;
    let endPage = currentPage + middlePage - 1;

    if (currentPage <= middlePage) {
      startPage = 1;
      endPage = maxPagesToShow;
    } else if (currentPage >= totalPages - middlePage + 1) {
      startPage = totalPages - maxPagesToShow + 1;
      endPage = totalPages;
    }

    for (let i = startPage; i <= endPage; i++) {
      pageNumbers.push(
        <span
          className={currentPage === i ? "selectedPage" : ""}
          onClick={(e) => selectPagehandler(e, i)}
          key={i}
        >
          {i}
        </span>
      );
    }
  }

  return pageNumbers;
};

return (<>
    {props?.selectedData?.length > 0 ?  (
      <div
        className="pagination"
        style={{ display: "flex", justifyContent: "center",columnGap: "2vw",paddingBottom:"4vh" }}
      >
        <div style={{display:"flex",columnGap: "0.5vw"}}>
          <Tooltip content="First">
            <Button
              size="small"
              appearance="subtle"
              disabled={props?.pages===firstPage?true:false}
              icon={<ArrowPrevious24Regular/>}
              onClick={()=>{
                  props.setPages(firstPage)
                  }}>
                  
              </Button>
          </Tooltip>
          
            <Tooltip content="Prev"> 
              <Button
              size="small"
                disabled={props?.pages <= 1}
                onClick={(e) => {
                  selectPagehandler(e, props?.pages - 1);
                }}
                appearance="subtle"
                iconPosition="after"
                icon={<IosArrowLtr24Regular/>}
              />
            </Tooltip>  
        </div>
        <div className="PageIndex" style={{ display: "flex"}}>
        <div style={{height:"95%",display:"flex",alignItems:"center",paddingRight:"1vw"}}>
        <Tooltip content="Select No. Of Task" relationship="label">
        <Dropdown size="small" 
        value={props?.numberOfTask}
        onOptionSelect={(e,data)=>{
            handleTaskNumber(data.optionValue)}}
        >
            {NumberTaskArray.map((option)=>{
                return(
                    <Option key={option} disabled={option === "Ferret"} text={option.toString()}>
                        {option}
                    </Option>
                )
                
            })}
          </Dropdown>
        </Tooltip>
         
          </div>
          <>
          {renderPageNumbers()}
          </>
        </div>
            <div style={{display:"flex",columnGap: "0.5vw"}}>
              <Tooltip content="Next">
              <Button
                  size="small"
                  disabled={Math.ceil(props?.selectedData.length / props?.numberOfTask) <= props?.pages}
                  onClick={(e) => {
                      selectPagehandler(e, props?.pages + 1);
                  }}
                  appearance="subtle"
                  iconPosition="center"
                  icon={<IosArrowRtl24Regular/>}
                />

              </Tooltip>
                <Tooltip content="Last">
                  <Button
                  disabled={props?.pages===lastPage?true:false}
                  appearance="subtle"
                  size="small"
                  onClick={()=>{props?.setPages(lastPage)}}
                  icon={<ArrowNext24Regular/>}
                  >
                  </Button>
                </Tooltip>  
            </div>
      </div>
    ):(<div style={{display:"flex",justifyContent:"center",height:"65%",alignItems:"center"}}>
      <Text size={500} weight="semibold" >No data available</Text>
      </div>)}

    </>
)

}
export default Pagination