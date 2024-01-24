import { useContext, useState, useEffect } from "react";
import { TeamsFxContext } from "./Context";
import AddTask from "./AddTask";
import { GetSite } from "./util";
import { app } from "@microsoft/teams-js";
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import {
  makeStyles,
  shorthands,
  TabList,
  Tab,
  Spinner,
  Card,
  CardHeader,
  Text,
  Persona,
  CardPreview,
  Badge,
} from "@fluentui/react-components";
import React from "react";
import { SearchBox } from "@fluentui/react-search-preview";
import OnGoing from "./tabListFile/OnGoing";
import UpComing from "./tabListFile/Upcoming";
import Completed from "./tabListFile/Completed";
import { useData } from "@microsoft/teamsfx-react";
import { getprofile ,getListDataCall} from "./util";

const useStyles = makeStyles({
  root: {
    alignItems: "flex-start",
    display: "flex",
    flexDirection: "column",
    justifyContent: "flex-start",
    ...shorthands.padding("50px", "20px"),
    rowGap: "20px",
  },
  textColor: {
    color: "white",
  },
});
export default function Tab1(props) {
  const { theme, themeString, teamsUserCredential } =
    useContext(TeamsFxContext);
  const [siteId, setSiteId] = useState("");
  const [listToDoId, setListToDo] = useState("");
  const [listToTaskEntryId, setListToDoTaskEntry] = useState("");
  const [needConsent, setNeedConsent] = useState(false);
  const [callReload, setCallReload] = useState(false);
  const [userData, setUserData] = useState([]);
  const [checkData, setCheckData] = useState(true);
  const [listTimeArry, setListTimeArry] = useState([]);
  const [upComingData, setUpComingData] = useState([]);
  const [selectedValue, setSelectedValue] = useState("OnGoing");
  const [loginuser, setLoginUser] = useState("");
  const [listData, setListData] = useState([]);
  const [userName, setUserName] = useState("");
  const [latestTaskData, setLatestTaskData] = useState([]);
  const [finalData, setFinalData] = useState([]);
  const [presence, setPresence] = useState("");
  const [addPermission, setAddPermissions] = useState(true);
  const[dialogVisbility,setDialogVisibility]=useState(false)
  const [countTask, setCounttask] = useState({
    CountOnGoing: 0,
    CountUpcoming: 0,
    CountCompleted: 0,
  });
  const settingSlider = {
    dots: true,
    infinite: false,
    speed: 500,
    slidesToShow: 1,
    slidesToScroll: 1,
  };
  useEffect(() => {
    if (teamsUserCredential) {
      getProfileIN(teamsUserCredential);
    }
  }, [teamsUserCredential]);
  const getProfileIN = async (teamsUserCredential) => {
    const tempPresence = await getprofile(teamsUserCredential);
    setPresence(tempPresence?.graphClientMessage?.availability);
  };

  const styles = useStyles();

  const { loading, data, reload } = useData(async () => {
    setCheckData(true);
    app.initialize().then(() => {
      app.getContext().then(async (context) => {
        let siteList;
        if (teamsUserCredential) {
          const userDispayName = await teamsUserCredential?.getUserInfo();
          const loginInfo = context.user;
          setUserName(userDispayName?.displayName);
          // const tanentUrl = context.sharePointSite.teamSiteDomain;
          setLoginUser(context.user);
          if (localStorage.getItem("getSiteList")) {
            siteList = JSON.parse(localStorage.getItem("getSiteList"));
            setSiteId(siteList.graphSiteid);
            setListToDo(siteList.graphListToDoId);
            setListToDoTaskEntry(siteList.graphListToTaskEntryId);
          } else {
            const obj = {
              siteName: "Teams_Site",
              listTodo: "ToDoTask",
              listTaskEntry: "To Do Task Entry",
              tanentUrl: context.sharePointSite.teamSiteDomain,
            };
            const res = await GetSite(teamsUserCredential, obj);
            const siteObj = {
              graphSiteid: res?.graphClientMessage,
              graphListToDoId: res?.listIdToDo,
              graphListToTaskEntryId: res?.listIdToDoEntry,
            };
            setSiteId(res?.graphClientMessage);
            setListToDo(res?.listIdToDo);
            setListToDoTaskEntry(res?.listIdToDoEntry);
            const siteObjString = JSON.stringify(siteObj);
            localStorage.setItem("getSiteList", siteObjString);
          }

          if (!teamsUserCredential) {
            throw new Error("TeamsFx SDK is not initialized.");
          }
          if (needConsent) {
            await teamsUserCredential.login(["User.Read"]);
            setNeedConsent(false);
          }
          if (localStorage.getItem("getSiteList")) {
            try {
              const obj = {
                siteId: JSON.parse(localStorage.getItem("getSiteList"))
                  .graphSiteid,
                listid1: JSON.parse(localStorage.getItem("getSiteList"))
                  .graphListToDoId,
                listid2: JSON.parse(localStorage.getItem("getSiteList"))
                  .graphListToTaskEntryId,
                userKey:"userProfile"
              };
              const functionRes = await getListDataCall(teamsUserCredential, obj);
              if (functionRes.NoUser === "No User Permissions") {
                setCheckData(false);
                setAddPermissions(false);
              } else if (
                functionRes.userInfo &&
                functionRes.graphClientMessage
              ) {
                setListData([]);
                setCounttask({
                  CountOnGoing: 0,
                  CountUpcoming: 0,
                  CountCompleted: 0,
                });
                setUpComingData([]);

                setUserData(functionRes?.userInfo?.value);
                functionRes?.graphClientMessage?.value
                  ?.sort((a, b) => {
                    return (
                      new Date(b.lastModifiedDateTime) -
                      new Date(a.lastModifiedDateTime)
                    );
                  })
                  .forEach((val) => {
                    if (
                      val.createdBy.user?.email ===
                        loginInfo?.userPrincipalName ||
                      val.fields?.ReviewerMail === loginInfo?.userPrincipalName
                    ) {
                      setListData((prev) => [...prev, val]);
                      if (
                        new Date(val.fields?.StartDate) <= new Date() &&
                        val.fields.Status !== "Completed"
                      ) {
                        setCounttask((prevObj) => ({
                          ...prevObj,
                          CountOnGoing: prevObj["CountOnGoing"] + 1,
                        }));
                      }
                      if (
                        val?.fields.Status !== "Completed" &&
                        new Date(val.fields?.StartDate) > new Date()
                      ) {
                        setCounttask((prevObj) => ({
                          ...prevObj,
                          CountUpcoming: prevObj["CountUpcoming"] + 1,
                        }));
                        setUpComingData((prev) => [...prev, val]);
                      }
                    }
                  });
                setCheckData(false);
                return functionRes.graphClientMessage.value;
              }
              setCheckData(false);
            } catch (error) {
              if (
                error.message.includes("The application may not be authorized.")
              ) {
                setNeedConsent(true);
              }
            }
          }
        }
      });
    });
  });
  const onTabSelect = (event, data) => {
    event.preventDefault();
    setSelectedValue(data.value);
  };
  const setSearch = (e) => {
    e.preventDefault();
    let newArry = listData.filter((item) => {
      return (
        item?.fields?.Title?.toLowerCase().includes(
          e.target.value.toLowerCase().trim()
        ) ||
        item.createdBy?.user.displayName
          ?.toLowerCase()
          .includes(e.target.value.toLowerCase().trim())
      );
    });
    if (newArry.length > 0) {
      setFinalData([]);
      setFinalData((prev) => [...prev, ...newArry]);
    } else {
      setFinalData([]);
    }
  
  };
  const handleLatestTask = (value) => {
    setLatestTaskData([]);
    setSelectedValue("latestTask");
    setLatestTaskData((prev) => [...prev, value]);
  };
  const checkPresence = (presence) => {
    if (presence === "DoNotDisturb") {
      return "do-not-disturb";
    } else if (presence === "BeRightBack") {
      return "be-right-back";
    } else {
      return presence?.toLowerCase();
    }
  };
  if (callReload) {
    reload();
    setCallReload(false);
  }
  return (
    <TeamsFxContext.Provider
      value={{
        theme,
        themeString,
        teamsUserCredential,
        loginuser,
        loading,
        userData,
        listTimeArry,
        siteId,
        listToDoId,
        listToTaskEntryId,
      }}
    >
      <div
        className={
          themeString === "default"
            ? "light"
            : themeString === "dark"
            ? "dark"
            : "contrast"
        }
      >
        <div>
          <Card className="CardProfile" style={{paddingLeft:"6px"}}>
            <div
              className="Main"
              style={{ display: "flex", justifyContent: "space-between"}}
            >
              <CardHeader
                // image={
                //   <img
                //     className={styles.logo}
                //     // src={resolveAsset("app_logo.svg")}
                //     alt="App name logo"
                //   />
                // }
                header={
                  <div>
                    <Persona
                      required
                      size="extra-large"
                      avatar={{
                        color: "colorful",
                        "aria-hidden": true,
                      }}
                      primaryText={
                        <Text >
                          <span className="profileName">{userName}</span>
                        </Text>
                      }
                      name={userName}
                      presence={{
                        status: checkPresence(presence),
                      }}
                      secondaryText={
                        <Text className={styles.textColor}>{presence}</Text>
                      }
                    />
                  </div>
                }
              />

              
                <CardPreview>
                  <div
                    className="headerTask"
                    style={{
                      paddingRight: "6vw",
                      paddingLeft: "3vw",
                      display: "flex",
                      flexDirection: "column",
                      justifyContent: "center",
                      rowGap: "0.5vw",
                    }}
                  >{!checkData&&(<>
                    <Text className={styles.textColor} size={300}>
                      Total task-{" "}
                      {countTask.CountOnGoing + countTask.CountUpcoming}
                    </Text>
                    <Text className={styles.textColor} size={300}>
                      Inprogress task- {countTask.CountOnGoing}
                    </Text>
                    <Text className={styles.textColor} size={300} style={{}}>
                      Upcoming task- {countTask.CountUpcoming}
                    </Text>
                  </>)}
                  </div>
                </CardPreview>
                <>
                  {upComingData.length > 0 ? (
                    <div className="upComingHeadCard">
                      <Text size={300} className={styles.textColor}>
                        Latest upcoming task
                      </Text>
                      <Slider {...settingSlider}>
                        {upComingData
                          ?.sort((a, b) => {
                            return (
                              new Date(a.fields.StartDate) -
                              new Date(b.fields.StartDate)
                            );
                          }).slice(0,4)
                          .map((value, index) => {
                            const Enddate = new Date(value?.fields.EndDate);
                            const DisplayEndDate = `${Enddate.getDate()}/${Enddate.getMonth()+1}/${Enddate.getFullYear()}`;
                            return (
                              <div
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  justifyContent: "center",
                                  backgroundColor: "transparent",
                                }}
                              >
                                <Card className="latestUpcomingCard"
                                  style={{
                                    background: "transparent",
                                    cursor: "pointer",
                                    paddingTop:'4px',
                                    padding:'6px',
                                    minHeight:"-moz-fit-content"
                                  }}
                                  onClick={(e) => {
                                    handleLatestTask(value);
                                  }}
                                >
                                  <CardHeader
                                    header={
                                      <diV style={{paddingBottom:"5px"}}>
                                      <Text
                                      truncate
                                      wrap={false}
                                        className={styles.textColor}
                                        size={300}
                                        style={{display:"block"}}
                                      >
                                        {value.fields.Title}
                                      </Text>
                                      </diV>
                                    }
                                  />
                                  <CardPreview>
                                    <div>
                                      <div
                                        className="upComingBody"
                                        style={{ display: "flex", }}
                                      >
                                        <Badge
                                          size="extra-large"
                                          shape="rounded"
                                          color="informative"
                                          appearance="outline"
                                          style={{minHeight:"3vw"}}
                                        >
                                          <div >
                                            <div
                                              className="dateBadge"
                                              style={{ height: "50%" }}
                                            >
                                              <Text
                                                className={styles.textColor}
                                                size={30}
                                                weight="bold"
                                              >
                                                {new Date(
                                                  value.fields.StartDate
                                                ).getDate()}
                                              </Text>
                                            </div>
                                            <div className="monthBadges" style={{width:"50%"}}>
                                              <Text
                                                className={styles.textColor}
                                                size={30}
                                              >
                                                {new Date(
                                                  value.fields.StartDate
                                                ).toLocaleString("default", {
                                                  month: "short",
                                                })}
                                              </Text>
                                            </div>
                                          </div>
                                        </Badge>
                                        <div style={{ paddingLeft: "5px" }}>
                                          <Text
                          
                                            className={styles.textColor}
                                            size={200}
                                          >
                                            End date : {DisplayEndDate}
                                          </Text>
                                          <br />
                                          <Text
                            
                                            className={styles.textColor}
                                            size={200}
                                          >
                                            {" "}
                                            Reviewer:{" "}
                                             {value.fields.ReviewerDipalyName}
                                          </Text>
                                        </div>
                                      </div>
                                    </div>
                                  </CardPreview>
                                </Card>
                              </div>
                            );
                          })}
                      </Slider>
                    </div>
                  ) : (
                    <div>
                      
                    </div>
                  )}
                </>
            </div>
          </Card>
        </div>
        {loading || checkData ? (
          <div
            style={{
              display: "flex",
              height: "100%",
              width: "100%",
              justifyContent: "center",
            }}
          >
            <Spinner label="Data loading" labelPosition="below"></Spinner>
          </div>
        ) : (
          <>
            <div className={styles.root}>
              <div
                className="headerBar"
                style={{
                  display: "flex",
                  justifyContent: "space-evenly",
                  alignItems: "self-end",
                  width: "100%",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    width: "65%",
                    justifyContent: "flex-end",
                  }}
                >
                  <SearchBox
                    className="searchBoxfield"
                    placeholder="Search task by title"
                    style={{ width: "50%" }}
                    onChange={(e) => {
                      setSearch(e);
                    }}
                    
                  />
                </div>
                <div
                  style={{
                    display: "flex",
                    width: "35%",
                    justifyContent: "end",
                  }}
                >
                  {addPermission && (
                    <AddTask
                      setSelectedValue={setSelectedValue}
                      setCallReload={setCallReload}
                      userName={userName}
                      listData={listData}
                    />
               
                  )}
                  
                </div>
              </div>
            </div>
            <TabList
              selectedValue={selectedValue}
              onTabSelect={onTabSelect}
              size="large"
            >
              <Tab value="OnGoing">Ongoing</Tab>
              <Tab value="UpComing">Upcoming</Tab>
              <Tab value="Completed">Completed</Tab>
            </TabList>

            <div>
              {selectedValue === "OnGoing" && (
                <div>
                  <OnGoing
                   SearchData={finalData}
                    setCallReload={setCallReload}
                    setDialogVisibility={setDialogVisibility}
                    listData={finalData.length > 0 ? finalData : listData}
                  />
                </div>
              )}
              {selectedValue === "UpComing" && (
                <div>
                  <UpComing
                    listData={finalData.length > 0 ? finalData : listData}
                  />
                </div>
              )}
              {selectedValue === "Completed" && (
                <div>
                  <Completed
                    setCallReload={setCallReload}
                    listData={finalData.length > 0 ? finalData : listData}
                  />
                </div>
              )}
              {selectedValue === "latestTask" && (
                <div>
                  <UpComing listData={latestTaskData} />
                </div>
              )}
            </div>
            {/* <Dialog  className="dialogPopup" open={dialogVisbility}
             modalProps={{
              isBlocking: false,
              styles: { main: { marginLeft: 'auto', marginRight: 0, marginTop: 20 } },
            }}
            
            >
              <DialogSurface>
                  <DialogBody>
                      <DialogContent style={{paddingTop:"20px",border:"GrayText"}}>
                        <div  className="taskCompletedDialog"> <div><CheckmarkCircle32Regular style={{color:"green"}}/></div>
                        <Text size={500}> Task Completed</Text>
                        </div>
                      </DialogContent>
                      <DialogActions style={{paddingRight:"25px"}}>
                        
                        <DialogTrigger >
                          <Button onClick={()=>{setDialogVisibility(false)}}>Ok</Button>
                        </DialogTrigger>
                      </DialogActions>
                  </DialogBody> 
              </DialogSurface>    
            </Dialog> */}
            {/* {dialogVisbility&&(<Toast>
        <ToastTitle>Task Completed</ToastTitle>
      </Toast>)} */}
          </>
        )}
        
      </div>
    </TeamsFxContext.Provider>
  );
}
