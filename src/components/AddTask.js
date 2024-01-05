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
  Text
} from "@fluentui/react-components";

import { ClipboardTaskAdd24Filled } from "@fluentui/react-icons";
import { addYears } from "@fluentui/react-calendar-compat";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { useContext, useState } from "react";
import { TeamsFxContext } from "./Context";
import { addTasklist,Notifiy } from "./util";
import  config from "./sample/lib/config"


const useStyles = makeStyles({
  content: {
    display: "flex",
    flexDirection: "column",
    rowGap: "10px",
  },
});
const AddTask = (props) => {
  const { teamsUserCredential, userData, listToDoId, siteId,loginuser } =
    useContext(TeamsFxContext);
  const [taskTitle, setTaskTitle] = useState("");
  const [descripiton, setDiscription] = useState("");
  const [EstimatedHours, setEstimatedHours] = useState("");
  const [StartDate, setsatrtDate] = useState("");
  const [EndDate, setEnddate] = useState("");
  const [isButtonDiabled, setButtonDisabled] = useState(true);
  const [ReviwerDisplayName, setReviwerDisplayName] = useState("");
  const [ReviewerEmail, setReviewerEmail] = useState();
  const [addTaskLoad, setAddTaskLoad] = useState(false)
  const [ReviewerUserId, setReviewerId] = useState("")
  const [popUpDialog,setPopUpDialog]=useState(false)
  const [addTaskRes,setAddtaskRes]=useState("")
  const styles = useStyles();
  const formateDate = (date) => {
    const selectedDate = new Date(date); // pass in date param here
    const formattedDate = `${
      selectedDate.getMonth() + 1
    }/${selectedDate.getDate()}/${selectedDate.getFullYear()}`;

    return formattedDate;
  };
  const sendNotification = {
    siteId: siteId,
    reviewerUserId:ReviewerUserId ,
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
          content:` 
            **${taskTitle}** Task is Added By ${props?.userName}

             Start Date : ${formateDate(StartDate)}
             End Date : ${formateDate(EndDate)}
             Estimated Hours : ${EstimatedHours}
           `,
        },
        toRecipients: [
          {
            emailAddress: {
              address:ReviewerEmail,
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
        EstimatedHours:EstimatedHours,
        ReviewerDipalyName: ReviwerDisplayName,
        ReviewerId: ReviewerUserId,
        Status: "UpComing",
        Descriptions: descripiton,
        ReviewerMail: ReviewerEmail,
      },
    };
    const AddTaskRes=await addTasklist(teamsUserCredential,obj)
    console.log("This is a response from backend",AddTaskRes)
    if (AddTaskRes.graphClientMessage==="Task Added"){
      await Notifiy(teamsUserCredential,sendNotification)
      setAddTaskLoad(false)
      // ClearState()
      // alert("Task added succcessfully")
      setAddtaskRes("succeeded")
      setPopUpDialog(true)
      setTimeout(onCloseFunction,1000)
      
    }
    else{
      setAddtaskRes("Task added Failed")
      setPopUpDialog(true)
      setTimeout(onCloseFunction,1000)
      
    }
  };
  const today = new Date();
  const maxDate = addYears(today,1);
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
 
  return (
    <div className="dialog_form">
      <Dialog modalType="alert">
        <DialogTrigger disableButtonEnhancement>
          <Button
            appearance="primary"
            icon={<ClipboardTaskAdd24Filled />}
          >
            Add task
          </Button>
        </DialogTrigger>
        {!popUpDialog?
        <DialogSurface aria-describedby={undefined}>
          <form
            onSubmit={(e) => {
              handleSubmit(e);
            }}
          >
            {addTaskLoad ? (
              <Spinner label="Adding New Task" labelPosition="below"></Spinner>
            ) : (
              <>
                <DialogBody>
                  <DialogTitle>Add new task</DialogTitle>
                  <DialogContent className={styles.content}>
                    <Label required htmlFor={"Task_title"}>
                      Task title
                    </Label>
                    <Input
                      required
                      type="text"
                      id={"Task_title"}
                      onChange={(e) => {
                        if (e.target?.value?.length < 35) {
                          setTaskTitle(e.target.value);
                        }
                      }}
                      value={taskTitle}
                      placeholder="Enter task title"
                    />
                    <Label required htmlFor={"descripition"}>
                      Descriptions
                    </Label>
                    <Input
                      required
                      type="text"
                      id={"description"}
                      onChange={(e) => {
                        if (e.target?.value?.length < 225) {
                          setDiscription(e.target.value);
                        }
                      }}
                      value={descripiton}
                      placeholder="Enter task descriptions"
                    />

                    <Label required htmlFor={"dateTime"}>
                      Startdate
                    </Label>
                    <DatePicker
                      required
                      minDate={new Date()}
                      maxDate={maxDate}
                      placeholder="Select start date "
                      formatDate={onFormatDate}
                      value={StartDate}
                      onSelectDate={setsatrtDate}
                      allowTextInput
                      className={styles.inputControl}
                    />
                    <Label required htmlFor={"EndDateTime"}>
                      Enddate
                    </Label>
                    <DatePicker
                      required
                      minDate={StartDate}
                      maxDate={maxDate}
                      placeholder="Select end date"
                      formatDate={onFormatDate}
                      allowTextInput
                      className={styles.inputControl}
                      value={EndDate}
                      onSelectDate={setEnddate}
                    />
                    <label required htmlFor={"EstimatedHour"}>
                      Estimated time <span style={{ color: "red" }}> *</span>
                    </label>
                    <Input
                      required
                      type="time"
                      id={"EstimatedHour"}
                      onChange={(e) => {
                        // if (e.target?.value < 100 && e.target?.value>=0) {
                          setEstimatedHours(e.target.value);
                        //  }
                      }}
                      value={EstimatedHours}
                      placeholder="Enter estimated time"
                    />

                    <label required htmlFor="Reviewer">
                    Reviewer <span style={{ color: "red" }}> *</span>
                    </label>
                    <Dropdown
                      required
                      id={"Reviwer"}
                      aria-labelledby="Reviwer"
                      onOptionSelect={(e, data) => {
                       
                        setReviewerEmail(data?.optionValue?.userPrincipalName);
                        setReviwerDisplayName(data?.optionText);
                        setReviewerId(data?.optionValue?.id);
                        

                        setButtonDisabled(false);
                      }}
                      placeholder="Select Reviewer"
                    >
                      {userData?.map((userData) => {
                        if(!(userData.id===loginuser.id)){
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
                      <Button appearance="secondary" onClick={ClearState}>
                        Close
                      </Button>
                    </DialogTrigger>
                    <Button
                      type="submit"
                      appearance="primary"
                      disabled={isButtonDiabled}
                    >
                      Submit
                    </Button>
                  </DialogActions>
                </DialogBody>
              </>
            )}
          </form>
        </DialogSurface>
        :(
        <DialogSurface>
        <DialogBody>
            <DialogContent>
              <Text size={300} weight="bold" style={{textAlign:"center"}}>Adding Task {addTaskRes}</Text>
            </DialogContent>
            <DialogActions>
                    <DialogTrigger disableButtonEnhancement>
                      <Button appearance="secondary" onClick={onCloseFunction}>
                        Close
                      </Button>
                    </DialogTrigger>
                  </DialogActions>
          </DialogBody>
          </DialogSurface>
          )}
      </Dialog>
    </div>
  );
};
export default AddTask;
