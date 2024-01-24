import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  Button,
  Input,
  Label,
  makeStyles,
  Dropdown,
  Option,
  Persona,
  Spinner,
  Text,
  Field,
  useId,
  Toaster,
  useToastController,
  ToastTitle,
  Toast,
} from "@fluentui/react-components";

import { ClipboardTaskAdd24Filled,CheckmarkCircle32Regular,Dismiss24Regular,ClipboardError24Regular} from "@fluentui/react-icons";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { useContext, useState } from "react";
import { TeamsFxContext } from "./Context";
import { addTasklist, Notifiy } from "./util";
import config from "./sample/lib/config";

const useStyles = makeStyles({
  content: {
    display: "flex",
    flexDirection: "column",
    rowGap: "10px",
  },
});
const AddTask = (props) => {
  const { teamsUserCredential, userData, listToDoId, siteId, loginuser } =
    useContext(TeamsFxContext);
  const [taskTitle, setTaskTitle] = useState("");
  const [descripiton, setDiscription] = useState("");
  const [EstimatedHours, setEstimatedHours] = useState("");
  const [StartDate, setsatrtDate] = useState("");
  const [EndDate, setEnddate] = useState("");
  const [isButtonDiabled, setButtonDisabled] = useState(true);
  const [ReviwerDisplayName, setReviwerDisplayName] = useState("");
  const [ReviewerEmail, setReviewerEmail] = useState();
  const [addTaskLoad, setAddTaskLoad] = useState(false);
  const [ReviewerUserId, setReviewerId] = useState("");
  const [popUpDialog, setPopUpDialog] = useState(false);
  const [addTaskRes, setAddtaskRes] = useState(false);
  const [validatin,setvalidation]=useState(false)
  const [validationEnddate,setValidationEnddate]=useState(true)
  const [taskRes,setTaskRes]=useState("")
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const styles = useStyles();
  const notifyToast = (value,type) =>{
        dispatchToast(
          <Toast>
            <ToastTitle weight="bold" size={300}>{value}</ToastTitle>
          </Toast>,
          { position:"top-end",intent:type,timeout:5000 }
        );
    }
  const formateDate = (date) => {
    const selectedDate = new Date(date); // pass in date param here
    const formattedDate = `${
      selectedDate.getMonth() + 1
    }/${selectedDate.getDate()}/${selectedDate.getFullYear()}`;

    return formattedDate;
  };
  const sendNotification = {
    siteId: siteId,
    reviewerUserId: ReviewerUserId,
    sendActivityNotification: {
      topic: {
        source: "text",
        value: "Task Added",
        webUrl: `https://teams.microsoft.com/l/entity/${config.teamsAppId}/index`,
      },
      activityType: "taskAdded",
      previewText: {
        content: `You Are Added as Rewier for Task ${taskTitle} By ${props?.userName} `,
      },
      templateParameters: [
        // {
        //   name: "taskId",
        //   value: (props?.element?.fields?.id).toString(),
        // },
        {
          name: "taskName",
          value: taskTitle.toString(),
        },
      ],
    },
    sendMail: {
      message: {
        subject: `${taskTitle} - Task Assigned To You`,
        body: {
          contentType: "Text",
          content: ` 
            **${taskTitle}** Task is Added By ${props?.userName}

             Start Date : ${formateDate(StartDate)}
             End Date : ${formateDate(EndDate)}
             Estimated Hours : ${EstimatedHours}
           `,
        },
        toRecipients: [
          {
            emailAddress: {
              address: ReviewerEmail,
            },
          },
        ],
        // ccRecipients: [
        //   {
        //     emailAddress: {
        //       address: "danas@contoso.onmicrosoft.com",
        //     },
        //   },
        // ],
      },
      saveToSentItems: "false",
    },
  };
const correctionEstimatedHr=(EstimatedHours)=>{
  if(EstimatedHours.toString().includes(".")){
    const estimateTimeArr=EstimatedHours.toString().split(".")

    if (estimateTimeArr[1].length<=1){
      return `${estimateTimeArr[0]}.0${estimateTimeArr[1]}`
    }
    else{
      return EstimatedHours
    }
  }
  else {
    if(EstimatedHours.toString()[0]==='0'&&!(EstimatedHours.toString().includes("."))){
      return EstimatedHours.split('0')[1]
    }  
    return EstimatedHours
  }

}
  const handleSubmit = async (e) => {
    e.preventDefault();
    setAddTaskLoad(true);
      const obj = {
        listId: listToDoId,
        siteId: siteId,
        field: {
          StartDate: StartDate,
          EndDate: EndDate,
          Title: taskTitle,
          EstimatedHours: correctionEstimatedHr(EstimatedHours),
          ReviewerDipalyName: ReviwerDisplayName,
          ReviewerId: ReviewerUserId,
          Status: "UpComing",
          Descriptions: descripiton,
          ReviewerMail: ReviewerEmail,
        },
      };
      const  AddTaskRes = await addTasklist(teamsUserCredential, obj);
      
      if (AddTaskRes.graphClientMessage === "Task Added") {
        await Notifiy(teamsUserCredential, sendNotification);
        notifyToast("Task added successfully","success")
        setTaskRes(false); 
      } else {
        notifyToast( "Task is not added successfully","error");
        setTaskRes(false);  
      }
      if(new Date(StartDate)>new Date()){
        props?.setSelectedValue("UpComing")
      }
      else{
        props?.setSelectedValue("OnGoing")
      }
      setTaskRes(false);
      setTimeout(()=>{props?.setCallReload(true)},1000)
    }
  const onFormatDate = (date) => {
    return !date
      ? ""
      : `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  };

  const ClearState = () => {
    setTaskTitle("");
    setDiscription("");
    setEstimatedHours("");
    setsatrtDate("");
    setEnddate("");
    setButtonDisabled(true);
    setReviwerDisplayName("");
    setReviewerEmail("");
  };
  const onCloseFunction = () => {
    setPopUpDialog(false);
    props?.setCallReload(true);
  };
const reviwerFieldSet=(e,data)=>{
    setReviewerEmail(
      data?.optionValue?.userPrincipalName
    );
    setReviwerDisplayName(data?.optionText);
    setReviewerId(data?.optionValue?.id);
    setButtonDisabled(false);
}
const checkAfterDecimal=(val)=>{
    let decimalValue=val.toString().split(".")
    if(decimalValue[1]<=60){
      return true
    }else{
      return false
    }
}
const checkStringLength=(val)=>{
  if(val.toString().indexOf(".")===1){
    return 4
  }
  else if(val.toString().indexOf(".")===2){
    return 5
  }
  

}
const estimatedTimeField=(e)=>{
 
  if (e.target?.value < 100 && e.target?.value >= 0
     &&(e.target.value.includes(".")?(e.target.value.toString().length<=(checkStringLength(e.target.value))):true)
     &&!(e.target.value.includes("00"))
      ){
    if(e.target.value.toString().includes(".")?checkAfterDecimal(e.target.value):true)

      setEstimatedHours(e.target.value);
      if(e.target.value==0){
        setButtonDisabled(true);
      }else{
           setvalidation(false)
           if(ReviwerDisplayName==="" ){
            setButtonDisabled(true);
           }
           else{
            setButtonDisabled(false);
           }   
      }
    }
}
const buttonDiabled=()=>{
  if(taskTitle===""||descripiton===""||StartDate===""||EndDate===""||EstimatedHours==="" || EstimatedHours==0||ReviwerDisplayName===""){
    return true
  }
  else{

    return false
  }

}
  return (
    <div className="dialog_form">
      <Dialog modalType="alert" open={taskRes} >
        <DialogTrigger disableButtonEnhancement >
          <Button appearance="primary" icon={<ClipboardTaskAdd24Filled />} onClick={()=>setTaskRes(true)}>
            Add task
          </Button>
        </DialogTrigger>
        {!popUpDialog ? (
          <DialogSurface >
            <form
              onSubmit={(e) => {
                handleSubmit(e);
              }}
            >
              {addTaskLoad ? (
                <Spinner style={{paddingTop:"12px"}}
                  label="Adding New Task"
                  labelPosition="below"
                ></Spinner>
              ) : (
                <>
                  <DialogBody style={{paddingRight:"12px"}}>
                    <DialogTitle
                        action={
                          <DialogTrigger action="close">
                            <Button
                              appearance="subtle"
                              aria-label="close"
                              icon={<Dismiss24Regular onClick={()=>{ClearState();setTaskRes(false)}} />}
                            />
                          </DialogTrigger>
                        }
                    >Add new task</DialogTitle>
                    <DialogContent className={styles.content}>
                      <Label required htmlFor={"Task_title"}>
                        Task title
                      </Label>
                      <Input
                        required
                        type="text"
                        id={"Task_title"}
                        onChange={(e) => {
                          if (
                            e.target?.value?.length < 35 && !(e.target.value.includes("  "))
                            
                          ) {
                            setTaskTitle(e.target.value.trimStart());
                          }
                        }}
                        value={taskTitle}
                        placeholder="Enter task title"
                      />
                      <Label required htmlFor={"descripition"}>
                        Description 
                      </Label>
                      <Input
                        required
                        type="text"
                        id={"description"}
                        onChange={(e) => {
                          if (e.target?.value?.length < 225 &&!(e.target.value.includes("  "))) {
                            setDiscription(e.target.value.trimStart());
                          }
                        }}
                        value={descripiton}
                        placeholder="Enter task descriptions"
                      />

                      <Field required htmlFor={"dateTime"}>
                        Start date
                      </Field>
                      <DatePicker
                        id="dateTime"
                        required
                        minDate={new Date()}
                        
                        placeholder="Select start date "
                        formatDate={onFormatDate}
                        value={StartDate}
                        onSelectDate={(date)=>{setsatrtDate(date);
                          // console.log("This a date selected",date)
                          if(new Date(EndDate)<new Date(date)){setEnddate("")}}}
                        className={styles.inputControl}
                        // allowTextInput={StartDate===""?true:false}
                       
                        
                      />
                      <Label required htmlFor={"EndDateTime"}>
                        End date
                      </Label>
                      <DatePicker
                        required
                        disabled={(StartDate==="")}
                        minDate={StartDate || new Date()}
                        
                        placeholder="Select end date"
                        formatDate={onFormatDate}
                        className={styles.inputControl}
                        value={EndDate}
                        onSelectDate={setEnddate}
                      />
                      <label required htmlFor={"EstimatedHour"}>
                        Estimated time<span style={{ color: "red" }}> *</span>
                      </label>
                      <Input
                        required
                        type="text"
                        id={"EstimatedHour"}
                        onChange={(e) => {
                          
                          estimatedTimeField(e)
                        }}
                        value={EstimatedHours}
                        placeholder="Enter estimated time"
                      />
                    {validatin&&(<Text style={{color:"red"}}>Estimated time cant be 0</Text>)}

                      <label required htmlFor="Reviewer">
                        Reviewer <span style={{ color: "red" }}> *</span>
                      </label>
                      <Dropdown
                       disabled={validatin}
                        required
                        id={"Reviwer"}
                        aria-labelledby="Reviwer"
                        onOptionSelect={(e, data) => {
                          reviwerFieldSet(e,data)
                        }}
                        placeholder="Select reviewer"
                      >
                        {userData?.map((userData) => {
                          if (!(userData.id === loginuser.id)) {
                            return (
                              <Option
                                required
                                text={userData.displayName}
                                key={userData.id}
                                value={userData}
                              >
                                <Persona
                                  required
                                  avatar={{
                                    color: "colorful",
                                    "aria-hidden": true,
                                  }}
                                  name={userData.displayName}
                                  presence={{
                                    status: "available",
                                  }}
                                  secondaryText="Available"
                                />
                              </Option>
                            );
                          }
                        })}
                      </Dropdown>
                    </DialogContent>
                    <DialogActions>
                      <DialogTrigger disableButtonEnhancement>
                        <Button appearance="secondary" onClick={()=>{ClearState();setTaskRes(false)}}>
                          Close
                        </Button>
                      </DialogTrigger>
                      <Button
                        type="submit"
                        appearance="primary"
                        disabled={buttonDiabled()||isButtonDiabled}
                      >
                        Submit
                      </Button>
                    </DialogActions>
                  </DialogBody>
                </>
              )}
            </form>
          </DialogSurface>
        ) : (
          <DialogSurface>
            <DialogBody>
              <DialogTitle 
              action={
                <DialogTrigger action="close">
                      <Button
                        appearance="subtle"
                        aria-label="close"
                        icon={<Dismiss24Regular />}
                        onClick={onCloseFunction}
                      />
                  </DialogTrigger>
              }/>
              <DialogContent>
                <Text size={300} weight="bold" style={{ textAlign: "center" }}>
                  {addTaskRes==="Task added successfully"&&( <div  className="taskCompletedDialog"> <div><CheckmarkCircle32Regular style={{color:"green"}}/></div>
                        <Text size={500}>{addTaskRes}</Text>
                        </div>)}
                  {addTaskRes==="Failed to add task"&&( <div  className="taskCompletedDialog"> <div><ClipboardError24Regular style={{color:"red"}}/></div>
                  <Text size={500}>{addTaskRes}</Text>
                  </div>)}

                  {addTaskRes==="Task already exists"&&(<span>{addTaskRes}</span>)}
                   
                </Text>
              </DialogContent>
              {/* <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button appearance="secondary" onClick={onCloseFunction}>
                    Close
                  </Button>
                </DialogTrigger>
              </DialogActions> */}
            </DialogBody>
          </DialogSurface>
        )}
      </Dialog>
      <Toaster toasterId={toasterId} />
    </div>
  );
};
export default AddTask;
