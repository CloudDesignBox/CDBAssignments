import * as React from 'react';
import styles from './MyAssignments.module.scss';
import { IMyAssignmentsProps,IMyAssignmentsState,AssignmentData} from './IMyAssignmentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as cachingService from "../../../services/cachingService";
import * as helperFunctions from "../../../services/helperFunctions";
import AssignmentItemDivV2 from "./AssignmentItem";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClientV3, } from "@microsoft/sp-http";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
//theme
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export default class MyAssignments extends React.Component<IMyAssignmentsProps, IMyAssignmentsState> {

  constructor(props){
    super(props);
    this.state={
      assignments:[],
      teams:[],
      courses:[],
      student:false,
      refreshTime:this.helperFunctions.getTimeNow(),
      currentPage:1,
      errorCode:""
     };
     this.refreshData = this.refreshData.bind(this);
     this.updatePaging = this.updatePaging.bind(this);
  }

  private CDBcachingService: cachingService.cachingService = new cachingService.cachingService(7200000);
  private helperFunctions: helperFunctions.helperFunctions = new helperFunctions.helperFunctions();
  private loadedAssignments:boolean = false;
  private loadedTeams:boolean=false;
  private loadedCourses:boolean=false;
  private loadedUserType:boolean=false;
  private cacheCheck:boolean=false;

  private cacheKey():string{
    let locationString:string=window.location.href;
    if(window.location.href.toLowerCase().indexOf("/sites/")>-1){
      locationString=window.location.href.split("/sites/")[1];
    }
    return `CDBMyAssignments${this.props.context.pageContext.user.loginName}${locationString}`;
  }

  public componentDidMount(){
    this.setPagingStyle();
  }

  public componentDidUpdate(){
    this.setPagingStyle();
  }

  //call from refresh button
  private refreshData(){
    this.cacheCheck=true;
    let assignments=[];
    let teams=[];
    let courses=[];
    let student=false;
    //check
    this.CDBcachingService.removeCache(`${this.cacheKey()}User`);
    this.CDBcachingService.removeCache(`${this.cacheKey()}Teams`);
    this.CDBcachingService.removeCache(`${this.cacheKey()}Courses`);
    this.loadedAssignments=false;
    this.loadedTeams=false;
    this.loadedCourses=false;
    this.loadedUserType=false;
    this.CDBcachingService.removeCache(`${this.cacheKey()}Assignments`);
    console.log("cache cleared");
    //update state
    this.setState({
      assignments:assignments,
      teams:teams,
      courses:courses,
      student:student,
      refreshTime:"Loading",
      currentPage:1
    });
  }

  private checkCache(){
    this.cacheCheck=true;
    let assignments=[];
    let teams=[];
    let courses=[];
    let student=false;
    let refreshTime="";
    //check
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}User`)){
      console.info("cached user loaded");
      student=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}User`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Assignments`)){
      console.info("cached assignments loaded");
      this.loadedAssignments=true;
      assignments=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Assignments`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Teams`)){
      console.info("cached teams loaded");
      this.loadedTeams=true;
      teams=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Teams`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Courses`)){
      console.info("cached courses loaded");
      this.loadedCourses=true;
      courses=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Courses`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Time`)){
      console.info("cached time loaded");
      refreshTime=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Time`);
    }
      //update state
      this.setState({
        assignments:assignments,
        teams:teams,
        courses:courses,
        student:student,
        refreshTime:refreshTime
      });
  }


  private getUserSDSType(){
    this.props.context.msGraphClientFactory.getClient('3')
    .then((client: MSGraphClientV3) => {
      client
        .api(`/education/me`)
        .version("v1.0")
        .select("primaryRole")
        .get((err, res) => {
          if(res){
            let user:MicrosoftGraph.EducationUser= res;
            if (user.primaryRole =="student"){
              this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}User`,true);
              this.setState({
                student:true
              });
            }else{
              this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}User`,false);
            }

          }else if(err){
            console.log("graph error "+err);
            //not sure this works
            if(err.toString().indexOf("InteractionRequiredAuthError")!=-1){
              this.setState({
                errorCode:`Warning - To use this web part, your SharePoint admin must accept the graph API permissions in the SharePoint admin centre https://${window.location.host.split(".")[0]}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement. Please log a support ticket with Cloud Design Box if you need further assistance.`
              });
            }
          }
        });
      });
  }


  private getTeamsForMe(){
    //GET
    //then match up via ids in items
    this.props.context.msGraphClientFactory.getClient('3')
    .then((client2: MSGraphClientV3) => {
      client2
        .api(`/me/joinedTeams`)
        .select("displayName,id,isArchived")
        .version("v1.0")
        .get((err2, res2) => {
          if(res2){
            let teams:MicrosoftGraph.Team[]= res2.value;
            this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Teams`,teams);
              this.setState({
                teams:teams
              });
          }else if(err2){
            console.log("graph error "+err2);
            //not sure this works
            if(err2.toString().indexOf("InteractionRequiredAuthError")!=-1){
              this.setState({
                errorCode:`Warning - To use this web part, your SharePoint admin must accept the graph API permissions in the SharePoint admin centre https://${window.location.host.split(".")[0]}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement. Please log a support ticket with Cloud Design Box if you need further assistance.`
              });
            }
          }
        });
      });
  }

  private getCoursesForMe(){
    //GET
    //then match up via ids in items
    this.props.context.msGraphClientFactory.getClient('3')
    .then((client2: MSGraphClientV3) => {
      client2
        .api(`/education/me/classes`)
        .select("id,course,displayName")
        .version("v1.0")
        .top(999)
        .get((err2, res2) => {
          if(res2){
            let courses:MicrosoftGraph.EducationClass[]= res2.value;
            this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Courses`,courses);
              this.setState({
                courses:courses
              });
          }else if(err2){
            console.log("graph error "+err2);
            //not sure this works
            if(err2.toString().indexOf("InteractionRequiredAuthError")!=-1){
              this.setState({
                errorCode:`Warning - To use this web part, your SharePoint admin must accept the graph API permissions in the SharePoint admin centre https://${window.location.host.split(".")[0]}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement. Please log a support ticket with Cloud Design Box if you need further assistance.`
              });
            }
          }
        });
      });
  }

  private getAssignmentsForMe(){
    this.props.context.msGraphClientFactory.getClient('3')
    .then((client2: MSGraphClientV3) => {
      client2
        .api(`/education/me/assignments?$expand=submissions&$top=999`)
        .version("v1.0")
        .get((err2, res2) => {
          if(res2){
            let assignments:MicrosoftGraph.EducationAssignment[] = res2.value;
            let tempAsssignment:AssignmentData[]=this.state.assignments;
            assignments.forEach(assignment => {
              this.helperFunctions.reportDebug(`Assignment loaded ${assignment.displayName}`);
              tempAsssignment.push({
                assignment:assignment,
                studentSubmissionDateStatus:"",
                currentPage:1,
                subjectName:""
              });
            });
            this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Assignments`,tempAsssignment);
              this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Time`,this.helperFunctions.getTimeNow());
              this.setState({
                assignments:tempAsssignment,
                refreshTime:this.helperFunctions.getTimeNow()
              });
          }else if(err2){
            console.log("graph error "+err2);
            //not sure this works
            if(err2.toString().indexOf("InteractionRequiredAuthError")!=-1){
              this.setState({
                errorCode:`Warning - To use this web part, your SharePoint admin must accept the graph API permissions in the SharePoint admin centre https://${window.location.host.split(".")[0]}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement. Please log a support ticket with Cloud Design Box if you need further assistance.`
              });
            }
          }
        });
      });

  }

  private sortClassesArray(assignments:AssignmentData[]):AssignmentData[]{
    let sortedAndCleanedData:AssignmentData[]=[];
    //remove assignments invalid or old
    //student version
    if(this.state.student){
      assignments.forEach(assignment => {
        //check status is live
        if((assignment.assignment.status== "assigned" || assignment.assignment.status== "published")){
          //check data is in future
          // if(new Date(assignment.assignment.dueDateTime) > new Date() ){
            if(assignment.assignment.submissions[0].status != "returned" && assignment.assignment.submissions[0].status != "submitted"){
              if(new Date(assignment.assignment.dueDateTime) > new Date() ){
                console.log(`This assignment is due in the future ${assignment.assignment.dueDateTime} ${assignment.assignment.submissions[0].status}`);
                assignment.studentSubmissionDateStatus="current";
              }else{
                console.log(`This assignment is overdue in the future ${assignment.assignment.dueDateTime} ${assignment.assignment.submissions[0].status}`);
                assignment.studentSubmissionDateStatus="overdue";
              }
            sortedAndCleanedData.push(assignment);
          }
        }
      });
    }else{
      //teacher version
      assignments.forEach(assignment => {
        //check status is live
        this.helperFunctions.reportDebug(`Assignment ${assignment.assignment.displayName} id: ${assignment.assignment.id} classid: ${assignment.assignment.classId}`);
        if((assignment.assignment.status== "assigned" || assignment.assignment.status== "published")){
          //check data is in future
          if(new Date(assignment.assignment.dueDateTime) > new Date() ){
            this.helperFunctions.reportDebug(`This assignment is overdue in the future ${assignment.assignment.dueDateTime}`);
            assignment.studentSubmissionDateStatus="current";
            sortedAndCleanedData.push(assignment);
          }else{
            this.helperFunctions.reportDebug(`This assignment is in the past  ${assignment.assignment.dueDateTime}`);
          }
        }
      });
    }
    //sort by due date
    sortedAndCleanedData.sort((a,b)=>{
      // Turn your strings into dates, and then subtract them
      // to get a value that is either negative, positive, or zero.
      let numbera:number = (new Date(b.assignment.dueDateTime).getTime());
      let numberb:number = (new Date(a.assignment.dueDateTime).getTime());
      return numberb - numbera;
    });
    // console.log(assignments);
    return sortedAndCleanedData;
  }

  private updatePaging(e){
    e.preventDefault();
    let pageNumber:number = parseInt(e.target.getAttribute("data-pagenumber"));
    this.setState({
      currentPage:pageNumber
    });
    return;
  }

  private setPagingStyle(){
    let allElements = this.props.context.domElement.querySelectorAll(`.${styles.paging}`);
    let elementArray=[];
    if(allElements){
      elementArray = Array.from(allElements);
    }
    elementArray.forEach(element => {
      if(parseInt(element.getAttribute("data-pagenumber")) == this.state.currentPage){
        element.classList.add(styles.pagingSelected);
      }else{
        element.classList.remove(styles.pagingSelected);
      }
    });
  }


///RENDER------

  public render(): React.ReactElement<IMyAssignmentsProps> {

    //load theme
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    //check cache
    if(!this.cacheCheck){
      this.checkCache();
    }

    //only load assignments after cache check if assignment length is still empty
    if(this.cacheCheck && this.state.assignments.length < 1 && !this.loadedAssignments){
      this.loadedAssignments=true;
      this.getAssignmentsForMe();
    }

    if(this.cacheCheck && this.state.teams.length < 1 && !this.loadedTeams){
      this.loadedTeams=true;
      this.getTeamsForMe();
    }

    if(this.cacheCheck && this.state.courses.length < 1 && !this.loadedCourses){
      this.loadedCourses=true;
      this.getCoursesForMe();
    }

    if(this.cacheCheck && !this.loadedUserType){
      this.loadedUserType=true;
      this.getUserSDSType();
    }

    //get paging from props
    let pagingProp:number=10; //default if not set
    if(this.props.webPartProps.pagingValue){
      pagingProp=this.props.webPartProps.pagingValue;
    }

    const listv2 = [];
    let pageCounter:number=1;
    let currentPage:number=1;
    const pages=[<span className={styles.paging} onClick={(e) => this.updatePaging(e)} data-pagenumber={currentPage}>{currentPage.toString()}</span>];
    if(this.state.assignments.length > 0 && this.state.teams.length > 0 && this.state.courses.length > 0){
      let sortedAssignments:AssignmentData[]=this.sortClassesArray(this.state.assignments);
      sortedAssignments.forEach(assignment => {

        let teamName:string="";
        let archivedState:boolean=false;
        let courseName:string="";
        let subject:string="";

        //get team name from list of teams
        this.state.teams.forEach(team => {
          if(team.id == assignment.assignment.classId){
            teamName=team.displayName;
            archivedState=team.isArchived;
          }
        });

        //get course name from list of teams
        this.state.courses.forEach(course => {
          if(course.id == assignment.assignment.classId){
            courseName=course.course.displayName;
            subject=course.course.subject;
          }
        });

        //show all outstanding assignments or hide overdue if set in the web part props
        if((this.props.webPartProps.hideOverDue && assignment.studentSubmissionDateStatus != "overdue")||!this.props.webPartProps.hideOverDue){
          //only show archived teams
          if((this.props.webPartProps.showArchivedTeams && archivedState) || !archivedState){
            //need to remove archived teams
            if(this.state.currentPage==currentPage){
              listv2.push(<AssignmentItemDivV2 itemData={assignment} teamName={teamName} course={courseName} subject={subject} />);
            }
            //set paging counter
            if(pageCounter==pagingProp){
              pageCounter=1;
              currentPage++;
              pages.push(<span className={styles.paging} onClick={(e) => this.updatePaging(e)} data-pagenumber={currentPage}>{currentPage.toString()}</span>);
            }else{
              pageCounter++;
            }
          }
        }
      });
    }else{
      listv2.push("No assignments found");
    }

    let warning:string="";

    if(this.state.errorCode){
      warning+=this.state.errorCode;
    }

    //block view
    let viewhtml=(<div>{listv2}</div>);

    return (
      <div className={ styles.myAssignments } style={{backgroundColor: semanticColors.bodyBackground}}>
        <section id="cdb-my-assignments">
          <div className={styles.header}>My Assignments</div>
          <div className={styles.warningbox}>{warning}</div>
          <Icon iconName="Refresh" onClick={this.refreshData}  className={styles.refreshicon}/>
          <span className={styles.refreshtext}>&nbsp;&nbsp;Last updated: {this.state.refreshTime}</span>

          {/* show view */}
          {viewhtml}

          <div className={styles.pagingContainer}>
          Pages {pages}
          </div>
          <div className={styles.footer}>Powered by <a href="https://www.clouddesignbox.co.uk" target="_blank">Cloud Design Box</a></div>
        </section>
      </div>
    );
  }
}
