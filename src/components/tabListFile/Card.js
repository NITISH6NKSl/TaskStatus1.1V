import * as React from "react";
import { useContext, useEffect } from "react";
import { TeamsFxContext } from "../Context";
import config from "../sample/lib/config";
import {  RemoveTask } from "../util";
import {
  makeStyles,
  shorthands,
  Button,
  Body1Strong,
  Title3,
  Field,
  ProgressBar,
  Spinner,
  Tooltip,
  Dialog,
  DialogActions,
  DialogBody,
  DialogSurface,
  DialogContent,
  DialogTitle,
  DialogTrigger,
  Text,
  Divider,
  Body1,
  useId,
  Toaster,
  useToastController,
  ToastTitle,
  Toast,
} from "@fluentui/react-components";
import {
  PlayCircle24Regular,
  RecordStop24Regular,
  Info24Regular,
  Delete24Filled,
  Dismiss24Regular,
} from "@fluentui/react-icons";
import {
  Card,
  CardHeader,
  CardFooter,
  CardPreview,
} from "@fluentui/react-components";
import { useState } from "react";
import { Notifiy, Update, playPause } from "../util";
const useStyles = makeStyles({
  main: {
    display: "flex",
    flexWrap: "wrap",
    flexDirection: "column",
    columnGap: "16px",
    rowGap: "36px",
  },

  title: {
    ...shorthands.margin(0, 0, "8px"),
  },
  textColor: {
    color: "white",
  },

  card: {
    maxWidth: "100%",
    height: "fit-content",
    marginBottom: "25px",
    backgroundColor: "transparent",
  },
  text: {
    ...shorthands.margin(0),
  },
  btn: {
    size: "small",
  },
  container: {
    ...shorthands.margin("5px", "0px"),
    backgroundColor: "brown"
  },
  cardbodyText: {
    marginRight: "35px",
  },


});

const CardComponent = (props) => {
  const [isPlay, setplay] = useState((props?.element?.fields?.IsPlay)?props.element.fields?.IsPlay:"Play");
  // const [completeBtn,setCompleteBtn]=((props?.element?.fields?.IsPlay)?false:true)
  const [completeBtn, setCompleteBtn] = (props?.element?.fields?.IsPlay ? [false, () => {}] : [true, () => {}]);
  const [isActualHourSet, setIsActualHourSet] = useState("");
  const [load, setLoad] = useState(false);
  const [completeDisable, setCompleteDisable] = useState(true);
  const [newPlay, setNewPlay] = useState("");
  const [ActualTime, setActualTime] = useState(0);
  const [deleteDialog,setDeleteDialog]=useState(false);
  const[taskCompleted,setTaskCompleted]=useState(false)
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const [removButton,setRemoveButton]=useState(false)
  const {
    teamsUserCredential,
    loginuser,
    themeString,
    siteId,
    listToDoId,
    listToTaskEntryId,
  } = useContext(TeamsFxContext);
  useEffect(() => {
    if (load) {
      handleToggelBtn(newPlay);
    }
  }, [newPlay]);
  useEffect(() => {
    if (props.listTimeArry?.length >= 2) {
      setActualHr();
    }
  }, [props.listTimeArry]);
  const notifyToast = (value,type) =>{
   return (dispatchToast(
    <Toast>
      <ToastTitle>{value}</ToastTitle>
    </Toast>,
    { position:"top-end",intent:type,timeout:3000}
  )
   )
     
  }

  const setActualHr = () => {
    let timeEntryArr = [];
    let listTimeArrId = [];
    props.listTimeArry?.forEach((time) => {
      if (time?.fields?.Id0 === props?.element?.fields?.id) {
        timeEntryArr.push(time.fields?.EntryExitTime);
        listTimeArrId.push(time?.fields?.id);
        return time.fields?.EntryExitTime;
      }
    });
    timeEntryArr = timeEntryArr.sort((a, b) => new Date(a) - new Date(b));

    let actualHour = 0;
    let actualMinute = 0;
    for (let i = 0; i < timeEntryArr.length; i += 2) {
      if (timeEntryArr.length !== i + 1) {
        const timeDifference =
          new Date(timeEntryArr[i + 1]) - new Date(timeEntryArr[i]);
        const hours = Math.floor(timeDifference / (1000 * 60 * 60));
        const minutes = Math.floor(
          (timeDifference % (1000 * 60 * 60)) / (1000 * 60)
        );
        actualHour += hours;
        actualMinute +=
        (parseInt(minutes) < 10
            ? parseInt("0" + minutes)
            : parseInt(minutes));
      }
      console.log("this is a minute at evey interval",actualMinute)
      if (parseInt(actualMinute) > 60) {
        actualHour += Math.floor(parseInt(actualMinute) / 60);
        actualMinute = Math.floor(parseInt(actualMinute) % 60);
      }
    }
    const formattedResult = actualHour.toString().padStart(2, '0') + ":" + actualMinute.toString().padStart(2, '0');
    setActualTime(formattedResult);
    // setActualTime(actualHour.toString() + ":" + actualMinute.toString());
  };
  console.log("This is a actual Hour of particular task",props?.element.fields.Title,ActualTime)
  const check = {
    date: true,
    setEstimateTime: false,
    setActualTime: false,
    setTaskButton: false,
    setProgressBar: false,
    completeBtnVisbile: false,
    reviwer: false,
    ActualStartBtn: false,
    isCreater: false,
    removeBtn: false,
  };
  const formateDate = (date) => {
    const selectedDate = new Date(date); // pass in date param here
    const formattedDate = `${
      selectedDate.getMonth() + 1
    }/${selectedDate.getDate()}/${selectedDate.getFullYear()}`;

    return formattedDate;
  };
  const sendNotification = {
    siteId: siteId,
    listId: listToTaskEntryId,
    componentId: props.element.fields.id,
    reviewerUserId: props?.element?.fields?.ReviewerId,
    sendActivityNotification: {
      topic: {
        source: "text",
        value: "Task Completed",
        webUrl: `https://teams.microsoft.com/l/entity/${config.teamsAppId}/index`,
      },
      activityType: "taskCompleted",
      previewText: {
        content: `${props?.element?.fields?.Title} Task Completed`,
      },
      templateParameters: [
        {
          name: "taskName",
          value: props?.element?.fields?.Title?.toString(),
        },
      ],
    },
    sendMail: {
      message: {
        subject: `${props?.element?.fields?.Title} - Completed `,
        body: {
          contentType: "Text",
          content: `
          **${props?.element?.fields?.Title}** Task is Completed By  ${
            props?.element?.createdBy.user?.displayName
          }

            Assignee : ${props?.element?.createdBy.user?.displayName}
            Status : Completed
            Reviwer : ${props?.element?.fields?.ReviewerDipalyName}
            Start Date : ${formateDate(props?.element?.fields?.StartDate)}
            End Date : ${formateDate(props?.element?.fields?.EndDate)}
            Actual Start Date : ${formateDate(
              props?.element?.fields?.ActualStartDate
            )}
            Estimated Hours : ${props?.element?.fields?.EstimatedHours}
            Actual Hour : ${ActualTime}`,
        },
        toRecipients: [
          {
            emailAddress: {
              address: props?.element?.fields?.ReviewerMail,
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
    status:"Complete",
  };
  const estimatedTime=(time)=>{

      if(time.toString().includes(".")){
        const estimatedTimeArray=props.element.fields.EstimatedHours.toString().split(".")
        const estimatedMinutes = parseInt(estimatedTimeArray[0]) * 60 + parseInt(estimatedTimeArray[1]);
        return estimatedMinutes;
      }
      else{
        const estimatedMinutes=parseInt(time)*60
        return estimatedMinutes
      }
  }
  const Styles = useStyles();
  const progreessValue=()=>{
    if(ActualTime){
    const actualTimeArray =ActualTime?.toString().split(":");
    const actualMinutes = parseInt(actualTimeArray[0]) * 60 + parseInt(actualTimeArray[1]);
    const EstimatedMinutes =estimatedTime(props?.element?.fields?.EstimatedHours);
     return (actualMinutes /EstimatedMinutes)
    }
    else{
      return 0
    }
    

  }
  // const setProgreessClass=()=>{
  //   if ( progreessValue() <= 0.8) {
  //     return "ongoingProgressBar";
  //   } else if ( progreessValue() > 0.8 && progreessValue() <= 0.99) {
  //     return "almostCompleteProgressBar";
  //   } else if ( progreessValue() > 0.99) {
  //     return "overdueProgressBar";
  //   }
  // }
  const ProgressColor = () => {
    if ( progreessValue() <= 0.8) {
      return "success";
    } else if ( progreessValue() > 0.8 && progreessValue() <= 0.99) {
      return "warning";
    } else if ( progreessValue() > 0.99) {
      return "error";
    }
  };
  const setLocalStroage = (data) => {
    localStorage.setItem("IsPlayCheck", data);
  };

  const handleToggelBtn = async (handle) => {
    const loadOngoingTab=async()=>{
      await props.setTaskLoad(true)
      setLoad(false);
    }
    if (handle) {
      if (!completeDisable) {
        setCompleteDisable(true);
      }
      let objMain;
      if (
        props.element.fields.ActualStartDate === undefined &&
        isActualHourSet === ""
      ) {
        const ActualStartDate = new Date();
        objMain = {
          siteId: siteId,
          listId: listToDoId,
          itemsId: props?.element?.fields?.id,
          field: { IsPlay: "Pause", ActualStartDate: ActualStartDate },
        };
      } else {
        objMain = {
          siteId: siteId,
          listId: listToDoId,
          itemsId: props.element.fields.id,
          field: { IsPlay: "Pause" },
        };
      }
      Update(teamsUserCredential, objMain).then((Response) => {
        setIsActualHourSet(
          Response?.receivedHTTPRequestBody?.field?.ActualStartDate
        );
        setplay(Response?.receivedHTTPRequestBody?.field?.IsPlay);
        setLocalStroage(Response?.receivedHTTPRequestBody?.field?.IsPlay);
      });
      var date = new Date();
      const obj = {
        siteId: siteId,
        listId: listToTaskEntryId,
        field: {
          Title: props.element.fields.Title,
          EntryExitTime: date,
          Id0: props?.element?.fields?.id,
        },
      };
      playPause(teamsUserCredential, obj).then((response) => {
          loadOngoingTab()
      });
    } else {
      Update(teamsUserCredential, {
        siteId: siteId,
        listId: listToDoId,
        itemsId: props.element.fields.id,
        field: { IsPlay: "Play" },
      }).then((Response) => {
        setplay(Response?.receivedHTTPRequestBody?.field?.IsPlay);
      });
      var timePause = new Date();
      const obj = {
        siteId: siteId,
        listId: listToTaskEntryId,
        field: {
          Title: props.element.fields.Title,
          EntryExitTime: timePause,
          Id0: props?.element?.fields?.id,
        },
      };
      playPause(teamsUserCredential, obj).then(async (response) => {
        let onPause = await props?.callLogData(true);
        if (onPause === "Success") {
          loadOngoingTab()
        } else if (onPause === "Success") {
          loadOngoingTab()
        }
      });
    }
  };
  const handleCompleteBtn = async (e) => {
    setTaskCompleted(true)
    const obj = {
      siteId: siteId,
      listId: listToDoId,
      itemsId: props?.element.fields.id,
      field: {
        ActualHours: ActualTime,
        Status: "Completed",
      },
    };
    const completeRes = await Update(teamsUserCredential, obj);
     if (completeRes.graphClientMessage?.Status === "Completed") {
      notifyToast("Task completed Successfully ","success")
    } else {
      notifyToast("Task is not completed Successfully","error")
    }
     
    try {
      await Notifiy(teamsUserCredential, sendNotification);
    } catch {}
    // props.setOnComplete(false);
    setTaskCompleted(false)
    setTimeout(()=>{props.setCallReload(true)})  
  };
  const deleteTask = async (id) => {
    setRemoveButton(true);
    
    const obj = {
      siteId: siteId,
      listId: listToDoId,
      itemsId: id,
    };
     const deleteRes=await RemoveTask(teamsUserCredential, obj)
 
      if(deleteRes.successMessage==="Deleted"){
        
        notifyToast("Task deleted successfully","success")
      }
      else{
        notifyToast("Task is not deleted successfully","error")
      }
    setRemoveButton(false);
    setDeleteDialog(false);
    setTimeout(()=>{props.setCallReload(true)},1000);
  };
  switch (props?.tabName) {
    case "OnGoing":
      if (
        loginuser?.userPrincipalName === props.element.createdBy.user.email &&
        loginuser?.userPrincipalName === props.element.fields.ReviewerMail
      ) {
        check.reviwer = false;
      } else if (
        loginuser?.userPrincipalName === props.element.fields.ReviewerMail
      ) {
        check.reviwer = true;
      }
      check.setTaskButton = true;
      check.setProgressBar = true;
      check.completeBtnVisbile = true;
      check.date = false;
      break;
    case "UpComing":
      check.setEstimateTime = true;
      break;
    case "Completed":
      if (loginuser?.userPrincipalName === props.element.createdBy.user.email) {
        check.isCreater = true;
      } else {
        check.isCreater = false;
      }
      check.setEstimateTime = true;
      check.setActualTime = true;
      check.ActualStartBtn = true;

      break;
    default:
      break;
  }
  const trimDescription = (str, maxLen, separator = " ") => {
    if (str?.length <= maxLen) return str;
    return str?.substr(0, str?.lastIndexOf(separator, maxLen));
  };
  return (
    <div>
      <div className="cardCompo">
        <Card className={Styles.card}>
          <CardPreview></CardPreview>
          <div>
            <CardHeader
              header={
                <div
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    width: "100%",
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "row",
                      justifyContent: "space-between",
                      width: "100%",
                    }}
                  >
                    <div>
                      <Title3 className={Styles.title}>
                        {props?.element?.fields?.Title}
                      </Title3>
                    </div>

                    <div style={{ cursor: "pointer" }} className="cardCompo">
                      <Dialog modalType="alert">
                        <DialogTrigger disableButtonEnhancement>
                          <Tooltip
                            withArrow
                            content="Details"
                            relationship="label"
                          >
                            <Info24Regular />
                          </Tooltip>
                        </DialogTrigger>
                        <DialogSurface>
                          <DialogBody>
                            <DialogTitle
                              weight="bold"
                              action={
                                <DialogTrigger action="close">
                                  <Button
                                    appearance="subtle"
                                    aria-label="close"
                                    icon={<Dismiss24Regular />}
                                  />
                                </DialogTrigger>
                              }
                            >
                              {props?.element?.fields.Title}
                              <Divider appearance="subtel"></Divider>
                            </DialogTitle>
                            <DialogContent>
                              <div
                                style={{
                                  display: "flex",
                                  flexDirection: "column",
                                  rowGap: "10px",
                                }}
                              >
                                <Text weight="bold">
                                  Description:{" "}
                                  <Text>
                                    {trimDescription(
                                      props?.element?.fields?.Descriptions,
                                      50
                                    )}
                                  </Text>
                                </Text>
                                <Text weight="bold">
                                  Start date:{" "}
                                  <Text>
                                    {formateDate(
                                      props?.element.fields.StartDate
                                    )}
                                  </Text>
                                </Text>
                                <Text weight="bold">
                                  End date:{" "}
                                  <Text>
                                    {formateDate(
                                      props?.element?.fields.EndDate
                                    )}
                                  </Text>
                                </Text>
                                <Text weight="bold">
                                  Estimated time:
                                  <Text>
                                    {" "}
                                    {props?.element?.fields?.EstimatedHours &&
                                    (props?.element?.fields?.EstimatedHours).toString().includes(
                                      "."
                                    )
                                      ? (props?.element?.fields?.EstimatedHours).toString().replace(
                                          ".",
                                          ":"
                                        )
                                      : `${props?.element?.fields?.EstimatedHours}:00`}
                                  </Text>
                                </Text>
                                <Text weight="bold">
                                  Reviewer:{" "}
                                  <Text>
                                    {props?.element?.fields.ReviewerDipalyName}
                                  </Text>
                                </Text>
                              </div>
                            </DialogContent>
                            {/* <DialogActions>
                              <DialogTrigger disableButtonEnhancement>
                                <Button appearance="primary">Close</Button>
                              </DialogTrigger>
                            </DialogActions> */}
                          </DialogBody>
                        </DialogSurface>
                      </Dialog>
                    </div>
                  </div>
                  <div>
                    <Divider appearance="subtel"></Divider>
                  </div>
                </div>
              }
              description={
                props?.element?.fields?.Descriptions ? (
                  <div style={{ marginTop: "5px" }} className={props?.tabName}>
                    <Dialog className="btnDescription">
                      <DialogTrigger disableButtonEnhancement>
                        <Tooltip
                          withArrow
                          content="Description"
                          relationship="label"
                        >
                          <Text >
                            <Text className="description" truncate wrap={false}>
                              <strong>Description: </strong>
                              {props?.element?.fields?.Descriptions}
                            </Text>
                          </Text>
                        </Tooltip>
                      </DialogTrigger>
                      <DialogSurface className="cardCompo">
                        <DialogBody>
                          <DialogTitle
                            action={
                              <DialogTrigger action="close">
                                <Button
                                  appearance="subtle"
                                  aria-label="close"
                                  icon={<Dismiss24Regular />}
                                />
                              </DialogTrigger>
                            }
                          >
                            Task Description
                            <Divider appearance="subtel"></Divider>
                          </DialogTitle>
                          
                          <DialogContent>
                            <Body1 weight="bold" size={300}>
                              {props?.element?.fields?.Descriptions}
                            </Body1>
                          </DialogContent>
                          {/* <DialogActions>
                        <DialogTrigger disableButtonEnhancement>
                          <Button appearance="secondary">Close</Button>
                        </DialogTrigger>
                      </DialogActions> */}
                        </DialogBody>
                      </DialogSurface>
                    </Dialog>
                  </div>
                ) : null
              }
            />
          </div>
          <div className={props?.tabName}>
            <CardFooter
              className={
                props?.tabName === "OnGoing"
                  ? "cardFooterContent"
                  : "footerComplete"
              }
            >
              <div className="fotterContent" style={{ display: "contents" }}>
                {check.date && (
                  <>
                    <Body1Strong className={Styles.cardbodyText}>
                      Start date :{"  "}
                      <span className="textValue">
                        {formateDate(props?.element?.fields?.StartDate)}
                      </span>
                    </Body1Strong>
                    <Body1Strong className={Styles.cardbodyText}>
                      End date :{"  "}
                      <span className="textValue">
                        {formateDate(props?.element?.fields?.EndDate)}
                      </span>
                    </Body1Strong>
                  </>
                )}
                {check.setEstimateTime && (
                  <Body1Strong>
                    Estimated time :{"  "}
                    <span className="textValue">
                      {props?.element?.fields?.EstimatedHours &&
                      (props?.element?.fields?.EstimatedHours).toString().includes(
                        "."
                      )
                        ? (props?.element?.fields?.EstimatedHours).toString().replace(
                            ".",
                            ":"
                          )
                        : `${props?.element?.fields?.EstimatedHours}:00`}
                    </span>
                  </Body1Strong>
                )}
                {check.ActualStartBtn && (
                  <Body1Strong>
                    Actual satrt date:{" "}
                    <span className="textValue">
                      {formateDate(props?.element?.fields?.ActualStartDate)}
                    </span>
                  </Body1Strong>
                )}
                {check.setActualTime && (
                  <Body1Strong>
                    Actual time:{"  "}
                    <span className="textValue">
                      {props?.element?.fields?.ActualHours}
                    </span>
                  </Body1Strong>
                )}
                {check.isCreater && (
                  <>
                      <Tooltip content="Remove task" withArrow>
                        <Delete24Filled
                          style={{ cursor: "pointer" }}
                          appearance="transparent"
                          onClick={() => {setDeleteDialog(true)}}
                        />
                      </Tooltip>
                  </>
                )}
                {check.reviwer ? (
                  ""
                ) : (
                  <>
                    {load  ? (
                      <Spinner></Spinner>
                    ) : (
                      <>
                        <div style={{ display: "flex", alignItems: "center" }}>
                          {isPlay &&
                            check.completeBtnVisbile &&
                            (isPlay === "Play" ? (
                              <Tooltip
                                withArrow
                                content="Start"
                                relationship="label"
                              >
                                <Button
                                  icon={<PlayCircle24Regular />}
                                  size="large"
                                  appearance="transparent"
                                  shape="circular"
                                  onClick={() => {
                                    setLoad(true);
                                    setNewPlay(true);
                                  }}
                                />
                              </Tooltip>
                            ) : (
                              <Tooltip
                                withArrow
                                content="Stop"
                                relationship="label"
                              >
                                <Button
                                  icon={<RecordStop24Regular />}
                                  size="large"
                                  shape="circular"
                                  appearance="transparent"
                                  onClick={() => {
                                    setLoad(true);
                                    setNewPlay(false);
                                  }}
                                />
                              </Tooltip>
                            ))}
                          {check.completeBtnVisbile && (
                            <Tooltip
                              withArrow
                              content="Complete task"
                              relationship="label"
                            >
                              <Button
                                className={themeString==='contrast'&&(isPlay === "Pause" || !completeDisable ||completeBtn)?"highContrast":""}
                                disabled={
                                  isPlay === "Pause" || !completeDisable ||completeBtn
                                }
                                appearance="primary"
                                onClick={(e)=>{handleCompleteBtn(e)}}
                                style={{ marginLeft: "1vw",}}
                              >
                                Completed
                              </Button>
                            </Tooltip>
                          )}
                        </div>
                      </>
                    )}
                  </>
                )}
              </div>
            </CardFooter>
          </div>
          {check.setProgressBar && (
            <div className="progressBar">
              <Field>
                <ProgressBar
                className="progress"
                  // style={{backgroundColor:"red"}}
                  thickness="large"
                  value={progreessValue()}
                  color={ProgressColor()}
                  
                />
                Task progressbar
              </Field>
            </div>
          )}
        </Card>
        <Dialog  open={deleteDialog}
          modalType="alert"
            >
              <DialogSurface>
                {!removButton?(
                  <DialogBody>
                  <DialogTitle 
                    action={
                      <DialogTrigger action="close">
                        <Button
                          appearance="subtle"
                          aria-label="close"
                          icon={<Dismiss24Regular onClick={()=>{setDeleteDialog(false)}} />}
                        />
                      </DialogTrigger>
                    }
                   />
                   
                    <DialogContent style={{display:"flex",justifyContent:"center",flexDirection:"column"}}>
                      <Text size={500} align="center"> Plaese confirm!</Text>
                      <Text size={300} align="center" style={{paddingTop:"18px",marginBottom:"5px"}}>Are you sure you want to delete this task permanently?</Text>
                    </DialogContent>
                  {/* <div> */}
                    <DialogActions style={{paddingRight:"171px"}}  >
                    
                      <DialogTrigger disableButtonEnhancement>
                        <Button onClick={()=>{deleteTask(props?.element?.fields?.id)}}
                        appearance="primary"
                        >
                        Delete</Button>
                      </DialogTrigger>
                      <DialogTrigger disableButtonEnhancement>
                        <Button onClick={()=>{setDeleteDialog(false)}}>Cancel</Button>
                      </DialogTrigger>
                    
                    </DialogActions>
                    {/* </div>  */}
                </DialogBody> 
                ):
                (<Spinner style={{paddingTop:"12px"}}
                  label="Removing Task"
                  labelPosition="below"
                ></Spinner>)
                }
                  
              </DialogSurface>    
            </Dialog>
            <Toaster toasterId={toasterId} />
            {taskCompleted&&(
              <Dialog open={taskCompleted}>
                <DialogSurface style={{paddingTop:"12px"}}>
                <Spinner label="Completing The Task" labelPosition="below" /> 
                  </DialogSurface>
              </Dialog>
            )}

      </div>
    </div>
  );
};
export default CardComponent;
