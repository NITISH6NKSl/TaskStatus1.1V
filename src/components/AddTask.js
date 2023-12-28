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
        IsPlay: "Play",
        EstimatedHours: EstimatedHours,
        ReviewerDipalyName: ReviwerDisplayName,
        ReviewerId: ReviewerUserId,
        Status: "UpComing",
        Descriptions: descripiton,
        ReviewerMail: ReviewerEmail,
      },
    };
    await addTasklist(teamsUserCredential, obj);
    await Notifiy(teamsUserCredential,sendNotification)
    props?.setCallReload(true);
  };
  const today = new Date();
  const maxDate = addYears(today, 1);

  const onFormatDate = (date) => {
    return !date
      ? ""
      : `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
  };

  const ClearState = (e) => {
    
    setTaskTitle("");
    setDiscription("");
    setEstimatedHours("");
    setsatrtDate("");
    setEnddate("");
    setButtonDisabled(true);
    setReviwerDisplayName("");
    setReviewerEmail("");
  };
  return (
    <div className="dialog_form">
      <Dialog modalType="alert">
        <DialogTrigger disableButtonEnhancement>
          <Button
            appearance="primary"
            icon={<ClipboardTaskAdd24Filled />}
            // onClick={(e) => {
            //   getUserData(e, teamsUserCredential);
            // }}
          >
            Add Task
          </Button>
        </DialogTrigger>

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
                  <DialogTitle>Add New Task</DialogTitle>
                  <DialogContent className={styles.content}>
                    <Label required htmlFor={"Task_title"}>
                      Task Title
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
                      placeholder="Enter a Task Title"
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
                      placeholder="Enter a Task Descriptions"
                    />

                    <Label required htmlFor={"dateTime"}>
                      Start Date Time
                    </Label>
                    <DatePicker
                      required
                      minDate={new Date()}
                      maxDate={maxDate}
                      placeholder="Select a date..."
                      formatDate={onFormatDate}
                      value={StartDate}
                      onSelectDate={setsatrtDate}
                      allowTextInput
                      className={styles.inputControl}
                    />
                    <Label required htmlFor={"EndDateTime"}>
                      End Date Time
                    </Label>
                    <DatePicker
                      required
                      minDate={StartDate}
                      maxDate={maxDate}
                      placeholder="Select a date..."
                      formatDate={onFormatDate}
                      allowTextInput
                      className={styles.inputControl}
                      value={EndDate}
                      onSelectDate={setEnddate}
                    />
                    <label required htmlFor={"EstimatedHour"}>
                      Estimated Hours <span style={{ color: "red" }}> *</span>
                    </label>
                    <Input
                      required
                      type="Number"
                      id={"EstimatedHour"}
                      onChange={(e) => {
                        if (e.target?.value < 100) {
                          setEstimatedHours(e.target.value);
                        }
                      }}
                      value={EstimatedHours}
                      placeholder="Enter a Estimated Hour in Number"
                    />

                    <label required htmlFor="Reviwer">
                      Reviwer <span style={{ color: "red" }}> *</span>
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
                      placeholder="Select Reviwer"
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
      </Dialog>
    </div>
  );
};
export default AddTask;
