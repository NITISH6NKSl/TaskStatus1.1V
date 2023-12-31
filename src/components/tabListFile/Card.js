import * as React from "react";
import { useContext, useEffect } from "react";
import { TeamsFxContext } from "../Context";
import config from "../sample/lib/config";
import { GetItems, RemoveTask } from "../util";
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
  Divider
} from "@fluentui/react-components";
import {
  PlayCircle24Regular,
  RecordStop24Regular,
  Info24Regular,
  Delete24Filled
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
  textColor:{
    color:"white"
  },

  card: {
    maxWidth: "100%",
    height: "fit-content",
    marginBottom: "25px",
    backgroundColor:"transparent",
  
  },
  text: {
    ...shorthands.margin(0),
  },
  btn: {
    size: "small",
  },
  container: {
    ...shorthands.margin("5px", "0px"),
  },
  cardbodyText: {
    marginRight: "35px",
  },
});

const CardComponent = (props) => {
  const [isPlay, setplay] = useState("");
  const [isActualHourSet, setIsActualHourSet] = useState("");
  const [load, setLoad] = useState(false);
  const [completeDisable,setCompleteDisable]=useState(true)
  const [loader, setLoader] = useState(true);
  const [newPlay, setNewPlay] = useState("");
  const[ActualTime,setActualTime]=useState(0)
  const {
    teamsUserCredential,
    loginuser,
    siteId,
    listToDoId,
    listToTaskEntryId,
  } = useContext(TeamsFxContext);
  const itemobj = {
    siteId: siteId,
    listToDoId: listToDoId,
    itemsId: props.element.fields.id,
  };
  useEffect(() => {
    if (load) {
      handleToggelBtn(newPlay);
    }
  },[newPlay]);
  useEffect(() => {
    if(props.tabName==="OnGoing"){
     
      GetItemsData(teamsUserCredential, itemobj);
    }
    else{
      setLoader(false);
    }
   
  },[props,teamsUserCredential]);
  useEffect(() => {
    if(props.listTimeArry?.length>=2){
      setActualHr()
    }
    
  }, [props.listTimeArry])
  const GetItemsData = async (teamsUserCredential, obj) => {
    const response = await GetItems(teamsUserCredential, obj);
    if (response?.fields?.IsPlay===undefined){
        setplay("Play")
        setCompleteDisable(false)
        setLoader(false)
        console.log("We are in if condition set play",response.fields.Title)
    }
    else{
      setplay(response?.fields?.IsPlay);
      setLoader(false);
      console.log("we are in a setplay else condition",response.fields.Title)
    }
    
    
  };
  const setActualHr=()=>{
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
        actualMinute += (Number(minutes) < 10) ? parseInt("0" + minutes, 10) : parseInt(minutes, 10);
      }
      if (Number(actualMinute) > 60) {
        actualHour += Math.floor(Number(actualMinute) / 60);
        actualMinute = Math.floor(Number(actualMinute) % 60);
      }
      
    }
    setActualTime(Number(actualHour.toString() +"."+ actualMinute.toString()));
  }
  const check = {
    date: true,
    setEstimateTime: false,
    setActualTime: false,
    setTaskButton: false,
    setProgressBar: false,
    completeBtnVisbile: false,
    reviwer: false,
    ActualStartBtn: false,
    isCreater:false,
    removeBtn:false,
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
          value: props?.element?.fields?.Title.toString(),
        },
      ],
    },
    sendMail: {
      message: {
        subject: `${props?.element?.fields?.Title} - Completed `,
        body: {
          contentType: "Text",
          content: `
          **${props?.element?.fields?.Title}** Task is Completed By  ${props?.element?.createdBy.user?.displayName}

            Assignee : ${props?.element?.createdBy.user?.displayName}
            Status : Completed
            Reviwer : ${props?.element?.fields?.ReviewerDipalyName}
            Start Date : ${formateDate(props?.element?.fields?.StartDate)}
            End Date : ${formateDate(props?.element?.fields?.EndDate)}
            Actual Start Date : ${formateDate(props?.element?.fields?.ActualStartDate)}
            Estimated Hours : ${(props?.element?.fields?.EstimatedHours)}
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
    status:"Complete"
  };
  const Styles = useStyles();
  let progreessValue =
    (ActualTime * 100) / (props.element.fields.EstimatedHours).toString().replace(":",".") / 100;
  const ProgressColor = () => {
    if (ActualTime && progreessValue <= 0.8) {
      return "success";
    } else if (ActualTime && progreessValue > 0.8 && progreessValue <= 0.99) {
      return "warning";
    } else if (ActualTime && progreessValue > 0.99) {
      return "error";
    }
  };
console.log("This is to check complete button of any task",props?.element?.fields?.Title,completeDisable)
  const setLocalStroage = (data) => {
    localStorage.setItem("IsPlayCheck", data);
  };

  const handleToggelBtn = (handle) => {

    if (handle) {
      if(!completeDisable){
        setCompleteDisable(true)
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
        setLoad(false);
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
      playPause(teamsUserCredential, obj).then((response) => {
        props.setOnPause(true)
        setLoad(false);
       
      });
    }
  };
  const handleCompleteBtn = async () => {
    props.setOnComplete(true)

    const obj = {
      siteId: siteId,
      listId: listToDoId,
      itemsId: props?.element.fields.id,
      field: {
        ActualHours: ActualTime,
        Status: "Completed",
      },
    };

    await Update(teamsUserCredential, obj);
    await props.setCallReload(true);
    try{
      await Notifiy(teamsUserCredential, sendNotification);
    }
    catch{
      
    }
    props.setOnComplete(false)
  };
  const deleteTask= async(id)=>{
    check.removeBtn=true
    const obj={
      siteId: siteId,
      listId: listToDoId,
      itemsId:id
    }
    await RemoveTask(teamsUserCredential,obj);
    check.removeBtn=false
    await props?.setCallReload(true);
  }
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
      if(loginuser?.userPrincipalName === props.element.createdBy.user.email){
        check.isCreater=true;
      }
      else{
        check.isCreater=false;
      }
      check.setEstimateTime = true;
      check.setActualTime = true;
      check.ActualStartBtn = true;
      
      break;
    default:
      break;
  }
  const trimDescription=(str, maxLen, separator = ' ')=>{
      if (str.length <= maxLen) return str;
      return str.substr(0, str.lastIndexOf(separator, maxLen));
  }
  return (
    <div >
      <div className="cardCompo" >
        <Card className={Styles.card} >
          <CardPreview></CardPreview>
          <div >
            <CardHeader
              header={
                <div style={{display:"flex",flexDirection:"row",justifyContent:"space-between",width:"100%", marginBottom:"5px"}}>
                  <div>
                  <Title3 className={Styles.title}>
                  {props?.element?.fields?.Title}
                  </Title3>
                  </div>
                  <div style={{cursor:"pointer"}} >
                      <Dialog modalType="alert">
                        <DialogTrigger disableButtonEnhancement>
                          <Tooltip     
                                withArrow
                                content="Details"
                                relationship="label"
                              >
                                <Info24Regular/>
                              </Tooltip>
                        </DialogTrigger>
                        <DialogSurface  className="cardCompo">
                          <DialogBody >
                            <DialogTitle weight="bold" action={null}>
                               {props?.element?.fields.Title}
                               <Divider appearance="strong" ></Divider>
                            </DialogTitle>
                            <DialogContent>
                            <div style={{display:"flex",flexDirection:"column",rowGap:"10px"}}>
                            
                                <Text weight="bold">Description: <Text>{trimDescription(props?.element?.fields?.Descriptions,50)}</Text></Text>
                                <Text weight="bold">
                                  Start date: <Text>{formateDate(props?.element.fields.StartDate)}</Text>
                                </Text>
                                <Text weight="bold">
                                  End date: <Text>{formateDate(props?.element?.fields.EndDate)}</Text>
                                </Text>
                                <Text weight="bold"><span>{props?.element?.fields?.EstimatedHours}</span>
                                  {/* Estimated time: <Text>  {props?.element?.fields?.EstimatedHours&&(props?.element?.fields?.EstimatedHours).toString().includes(".")?
                    (props?.element?.fields?.EstimatedHours).toString().replace(".",":"):
                    `${props?.element?.fields?.EstimatedHours}:00` */}
                    
                                </Text>
                                <Text weight="bold">
                                Reviewer: <Text>{props?.element?.fields.ReviewerDipalyName}</Text>
                                </Text>
                              </div>
                    
                            </DialogContent>
                            <DialogActions>
                              <DialogTrigger disableButtonEnhancement>
                                <Button appearance="primary">Close</Button>
                              </DialogTrigger>
                            </DialogActions>
                          </DialogBody>
                        </DialogSurface>
                      </Dialog>                
                  </div>
                </div>
              }
              description={
                props?.element?.fields?.Descriptions ? 
                <Dialog
                className="btnDescription"
                style={{ display: "flex" }}
              >
                  <DialogTrigger disableButtonEnhancement>
                      <Tooltip
                        withArrow
                        content="Description"
                        relationship="label"
                      >
                        <Body1Strong className="description"  wrap={true} >
                            Description: {trimDescription(props?.element?.fields?.Descriptions,60)}
                            <span></span>
                        </Body1Strong>
                      </Tooltip>
                  </DialogTrigger>
                  <DialogSurface className="cardCompo">
                    <DialogBody>
                      <DialogTitle>Task description</DialogTitle>
                      <DialogContent>
                        <Text weight="bold" size={300}>{props?.element?.fields?.Descriptions}</Text>
                      </DialogContent>
                      <DialogActions>
                        <DialogTrigger disableButtonEnhancement>
                          <Button appearance="secondary">Close</Button>
                        </DialogTrigger>
                      </DialogActions>
                    </DialogBody>
                  </DialogSurface>
              </Dialog>
                   : null    
              }
            />
          </div>
          <div style={{marginTop:"5px"}}>
          <CardFooter
           className={props?.tabName==="OnGoing"?("cardFooterContent"):"footerComplete"}
          >
            <div className="fotterContent" style={{ display: "contents"}}>
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
                 {props?.element?.fields?.EstimatedHours}
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
              {check.isCreater&&(<>
              {!check.removeBtn?
              <Tooltip content="Remove Task" withArrow>
              <Delete24Filled style={{cursor:"pointer"}} appearance="transparent" onClick={()=>deleteTask(props?.element?.fields?.id)}/>
              </Tooltip>
              :<Spinner label="Removing Task" labelPosition="below"></Spinner>
              }
              </>)}
              {check.reviwer ? (
                ""
              ) : (
                <>
                  {load || loader ? (
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
                              <Button icon={<PlayCircle24Regular/>}
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
                            <Button icon={<RecordStop24Regular/>}
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
                          <Tooltip  withArrow
                          content="Complete Task"
                          relationship="label">
                          <Button
                            disabled={isPlay === "Pause"||!completeDisable}
                            appearance="primary"
                            onClick={handleCompleteBtn}
                            style={{ marginLeft: "1vw" }}
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
              <Field
              >
                <ProgressBar
                  className={Styles.container}
                  thickness="large"
                  value={progreessValue}
                  color={ProgressColor()}
                />
                Task progressbar
              </Field>
            </div>
          )}
        </Card>
      </div>
    </div>
  );
};
export default CardComponent;
