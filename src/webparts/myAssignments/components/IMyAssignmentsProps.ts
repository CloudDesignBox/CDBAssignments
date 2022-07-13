import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  SPHttpClient   
} from '@microsoft/sp-http';
export interface IMyAssignmentsProps {
  context: WebPartContext;
  sphttpContext:SPHttpClient;
  webPartProps:IMyAssignmentsWebPartProps;
  themeVariant: IReadonlyTheme | undefined;
}

export interface IMyAssignmentsState {
  assignments:AssignmentData[];
  classes:CDBClassTeams[];
  student:boolean;
  refreshTime:string;
  currentPage:number;
  errorCode:string;
  filteredSubject:string;
}

export interface CDBConfig {
  ID:number;
  classesUrl:string;
  myClassesList:string;
  MultiSchoolPath:string;
}


export interface AssignmentData {
  assignment:MicrosoftGraph.EducationAssignment;
  studentSubmissionDateStatus:string;
  currentPage:number;
}

export interface CDBClassTeams {
  ID:number;
  Title:string;
  ClassMembers:string;
  ClassMembersEmail:string;
  ClassTeachers:string;
  GroupId:string;
  SubjectName:string;
  SubjectSiteResourcesUrl:string;
}

export interface IExampleItem {
  thumbnail: string;
  name: string;
  description: string;
  color: string;
  shape: string;
  location: string;
  width: number;
  height: number;
}
export interface IMyAssignmentsWebPartProps {
  pagingValue:number;
  subjectFilter:boolean;
  hideOverDue:boolean;
}



