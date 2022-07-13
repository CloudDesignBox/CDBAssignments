import * as React from 'react';
import styles from './MyAssignments.module.scss';
import { IMyAssignmentsProps,IMyAssignmentsState,AssignmentData,CDBConfig,CDBClassTeams} from './IMyAssignmentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as cachingService from "../../../services/cachingService";
import * as helperFunctions from "../../../services/helperFunctions";
import AssignmentItemDivV2 from "./AssignmentItemv2";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClient, SPHttpClient } from "@microsoft/sp-http";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
//theme 
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export default class MyAssignments extends React.Component<IMyAssignmentsProps, IMyAssignmentsState> {

  constructor(props){
    super(props);
    this.state={
      assignments:[],
      student:false,
      classes:[],
      config:null,
      refreshTime:this.helperFunctions.getTimeNow(),
      currentPage:1,
      errorCode:"",
      filteredSubject:""
     };

     this.refreshData = this.refreshData.bind(this);
     this.updatePaging = this.updatePaging.bind(this);
  }

  private CDBcachingService: cachingService.cachingService = new cachingService.cachingService(7200000);
  private helperFunctions: helperFunctions.helperFunctions = new helperFunctions.helperFunctions();
  private loadedAssignments:boolean = false;
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
    let config=null;
    let assignments=[];
    let classes=[];
    let student=false;
    
    //check cdbconfig
    this.CDBcachingService.removeCache(`${this.cacheKey()}Config`);
    this.CDBcachingService.removeCache(`${this.cacheKey()}User`);
    this.CDBcachingService.removeCache(`${this.cacheKey()}Classes`);
    this.loadedAssignments=false;
    this.CDBcachingService.removeCache(`${this.cacheKey()}Assignments`);
    this.CDBcachingService.removeCache(`${this.cacheKey()}filteredSubject`);
    console.log("cache cleared");
    //update state
    this.setState({
      assignments:assignments,
      classes:classes,
      config:config,
      student:student,
      refreshTime:"Loading",
      currentPage:1,
      filteredSubject:""
    });
  }

  
  private checkCache(){
    this.cacheCheck=true;
    let config=null;
    let assignments=[];
    let classes=[];
    let student=false;
    let refreshTime="";
    let filteredSubject:string="";
    //check cdbconfig
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}User`)){
      console.info("cached user loaded");
      student=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}User`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Assignments`)){
      console.info("cached assignments loaded");
      this.loadedAssignments=true;
      assignments=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Assignments`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Time`)){
      console.info("cached time loaded");
      refreshTime=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}Time`);
    }
    if(this.CDBcachingService.getWithExpiry(`${this.cacheKey()}filteredSubject`)){
      console.info("cached filteredSubject loaded");
      filteredSubject=this.CDBcachingService.getWithExpiry(`${this.cacheKey()}filteredSubject`);
    }
      //update state
      this.setState({
        assignments:assignments,
        classes:classes,
        config:config,
        student:student,
        refreshTime:refreshTime,
        filteredSubject:filteredSubject
      });

  }



  private  getAssignmentsAll(){
    let classCounter:number=0;
    this.state.classes.forEach(singleClass => {
      classCounter++;
      //limit assignment calls to first 25 classes
      if(classCounter<25){
        this.getAssignments(singleClass);
      }
    });
  }

  private getAssignments(singleClass:CDBClassTeams){
    //load assignments
    this.props.context.msGraphClientFactory.getClient()
    .then((client2: MSGraphClient) => {
      client2
        .api(`/education/classes/${singleClass.GroupId}/assignments?$expand=submissions&$top=999`)
        .version("v1.0")
        .get((err2, res2) => {
          if(res2){
            let assignments:MicrosoftGraph.EducationAssignment[] = res2.value;
            assignments.forEach(assignment => {
              let tempAsssignment:AssignmentData[]=this.state.assignments;
              tempAsssignment.push({
                assignment:assignment,
                teamData:singleClass,
                studentSubmissionDateStatus:"",
                currentPage:1
              });
              this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Assignments`,tempAsssignment);
              this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Time`,this.helperFunctions.getTimeNow());
              this.helperFunctions.reportDebug(`Assignment loaded ${assignment.displayName}`);
              this.setState({ 
                assignments:tempAsssignment,
                refreshTime:this.helperFunctions.getTimeNow()
              });
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
        this.helperFunctions.reportDebug(`Assignment ${assignment.assignment.displayName} id: ${assignment.assignment.id} classid: ${assignment.teamData.GroupId}`);
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



  
   private loadClassData(items: CDBClassTeams[]):void{
     let classItems: CDBClassTeams[]=[];
     let foundStudent:boolean = false;
     let foundTeacher:boolean = false;
     let studentFlag:boolean=false;
     let filteredSubject:string="";
     //check classes for student
    items.forEach((item: CDBClassTeams) => {
      foundTeacher=false;
      foundStudent=false;
      let excludeClass:boolean=false;
      if(item.ClassMembers){
        //check if user is a student with email
        if(item.ClassMembers.length>0){
          let studentemails=item.ClassMembersEmail.split(";");
          studentemails.forEach(studentemail => {
            if(studentemail.toLowerCase() == this.props.context.pageContext.user.email.toLowerCase()){
              foundStudent=true;
              studentFlag=true;
            }
          });
        }
        //check if user is a student with email
        if(item.ClassMembers.length>0){
          let studentusernames=item.ClassMembers.split(";");
          studentusernames.forEach(studentusername => {
            if(studentusername.toLowerCase() == this.props.context.pageContext.user.loginName.toLowerCase()){
              foundStudent=true;
              studentFlag=true;
            }
          });
        }   
      }
      //check for teacher
      if(item.ClassTeachers){
        //check if user is a student with email
        if(item.ClassTeachers.length>0){
          let teacheremails=item.ClassTeachers.split(";");
          teacheremails.forEach(teacheremail => {
            if(teacheremail.toLowerCase() == this.props.context.pageContext.user.email.toLowerCase() || teacheremail.toLowerCase() == this.props.context.pageContext.user.loginName.toLowerCase()){
              foundTeacher=true;
            }
          });
        }
      }

      //check to see if subject site filter exists
    if(this.props.webPartProps.subjectFilter){
      excludeClass=true;
      if(item.SubjectSiteResourcesUrl && item.SubjectName){
        if(window.location.href.toLowerCase().indexOf(item.SubjectSiteResourcesUrl.toLowerCase())>-1 && (item.SubjectSiteResourcesUrl.toLowerCase().indexOf("/sites/")>-1 || item.SubjectSiteResourcesUrl.toLowerCase().indexOf("/teams/")>-1 )){
          //extra check to filter out subject codes that start with another subject code e.g. En and ENG
          //either its the end of the url or the next char is a slash or question mark
          let nextChar:string=window.location.href.toLowerCase().substring(window.location.href.toLowerCase().indexOf(item.SubjectSiteResourcesUrl.toLowerCase())+item.SubjectSiteResourcesUrl.length,window.location.href.toLowerCase().indexOf(item.SubjectSiteResourcesUrl.toLowerCase())+1+item.SubjectSiteResourcesUrl.length);
          if(window.location.href.toLowerCase().endsWith(item.SubjectSiteResourcesUrl.toLowerCase()) || nextChar =="/" || nextChar=="?"){
            excludeClass=false;
            filteredSubject=item.SubjectName;
          }
        }
      }
    }


    if((foundStudent || foundTeacher)&&!excludeClass){
        classItems.push(item);
        this.helperFunctions.reportDebug(`Team loaded ${item.Title} as a teacher or student`);
      }else{
        this.helperFunctions.reportDebug(`Warning: Team found ${item.Title} but the logged in user was not listed as a student or teacher.`);
      }
      this.helperFunctions.reportDebug(`Student: ${foundStudent}. Teacher: ${foundTeacher}`);
      
    });
    
    this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}User`,studentFlag);
    this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}Teams`,classItems);
    this.CDBcachingService.setWithGlobalExpiry(`${this.cacheKey()}filteredSubject`,filteredSubject);
          
    //set state
    this.setState({ 
      classes:classItems,
      student:studentFlag,
      filteredSubject:filteredSubject
    });
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

    if(this.state.classes.length > 0 && !this.loadedAssignments){
      this.loadedAssignments=true;
      this.getAssignmentsAll();
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
    if(this.state.assignments.length > 0){
      let sortedAssignments:AssignmentData[]=this.sortClassesArray(this.state.assignments);
      sortedAssignments.forEach(assignment => {

        //show all outstanding assignments or hide overdue if set in the web part props
        if((this.props.webPartProps.hideOverDue && assignment.studentSubmissionDateStatus != "overdue")||!this.props.webPartProps.hideOverDue){

          //need to remove archived teams
          if(this.state.currentPage==currentPage){
            listv2.push(<AssignmentItemDivV2 assignment={assignment.assignment} teamData={assignment.teamData} studentSubmissionDateStatus={assignment.studentSubmissionDateStatus} currentPage={currentPage} />);
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
        

      });
    }else{
      listv2.push("No assignments found");
    }

    let warning:string="";
    if(this.state.classes.length>25){
      warning="Warning - we have detected that you are in over 25 classes. Due to throttling limits, this web part does not support more than 25 requests to graph and therefore it may not show a complete list of assignments.";
    }
    if(this.state.classes.length=0){
      warning="We have not detected any classes that you teach or are a studnet of in class dashboard.";
    }
    if(this.state.errorCode){
      warning+=this.state.errorCode;
    }
    if(this.state.classes.length>24){
      warning="Warning - we have detected that you are in over 24 classes. Due to throttling limits, this web part does not support more than 24 requests to graph and therefore it may not show a complete list of assignments. ";
    }
    //block view
    let viewhtml=(<div>{listv2}</div>);

    return (
      <div className={ styles.myAssignments } style={{backgroundColor: semanticColors.bodyBackground}}>
        <section id="cdb-my-assignments">
          <div className={styles.header}>My Assignments {this.state.filteredSubject && <span>for {this.state.filteredSubject}</span>}</div>
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













  // private getTeamsData(){
  //   this.loadedTeams = true;
  //   console.log("loading teams");
  //   let numberofClasses:number=0;
    
  //   this.props.context.msGraphClientFactory.getClient()
  //   .then((client: MSGraphClient) => {
  //     client
  //       .api(`/education/me/classes?$top=999`)
  //       .version("v1.0")
  //       .get((err, res) => {
  //         if(res){
  //             let classes:MicrosoftGraph.EducationClass[] = res.value;
  //             numberofClasses = classes.length;
  //             // classes.forEach(singleClass => {
  //             //   console.log(singleClass.displayName);
  //             //   // throtting with this so disabled
  //             //   this.getArchiveStatus(singleClass,numberofClasses);
  //             //   });
  //             // 
  //             this.setState({ 
  //               classes:classes
  //             });
  //           } 
  //         });
  //       });
  // }



  // private getArchiveStatus(singleClass:MicrosoftGraph.EducationClass,numberofClasses:number){
  //     console.log(`getting archive status of ${singleClass.displayName}`);
  //      //check archive status of class
  //      this.props.context.msGraphClientFactory.getClient()
  //      .then((client3: MSGraphClient) => {
  //        client3
  //          .api(`/teams/${singleClass.id}?$select=IsArchived`)
  //          .version("v1.0")
  //          .get((err3, res3) => {
  //            if(res3){
  //            let ClassTeam:MicrosoftGraph.Team = res3;
  //            this.classCount++;
             
  //            if(!ClassTeam.isArchived){
  //              console.log(`found unarchived class team ${singleClass.displayName}`);
  //              this.tempClasses.push(singleClass);
  //              //only update state when all classes archive state loaded
  //              if(this.classCount == numberofClasses){
  //                this.setState({ 
  //                  classes:this.tempClasses
  //                });
  //              }
               
  //            }else{
  //              console.log(`ignoring archived class team ${singleClass.displayName}`);
  //            }
  //          }else{
  //            console.log(`found unarchived class team ${singleClass.displayName}`);
  //              this.tempClasses.push(singleClass);
  //              //only update state when all classes archive state loaded
  //              if(this.classCount == numberofClasses){
  //                this.setState({ 
  //                  classes:this.tempClasses
  //                });
  //              }
  //          }
  //          });
  //        });
  // }


  

  // private getUserSDSType(){
  //   this.loadedUser=true;
  //   console.log("loading teams");
  //   this.props.context.msGraphClientFactory.getClient()
  //   .then((client: MSGraphClient) => {
  //     client
  //       .api(`/education/me`)
  //       .version("v1.0")
  //       .get((err, res) => {
  //         if(res){
  //             let user:MicrosoftGraph.EducationUser = res;
  //             if (user.primaryRole =="student"){
  //               this.setState({ 
  //                 user:user
  //               });
  //             }
  //         }
  //       });
  //     });

  // }
