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
  teams:MicrosoftGraph.Team[];
  courses:MicrosoftGraph.EducationClass[];
  student:boolean;
  refreshTime:string;
  currentPage:number;
  errorCode:string;
}

export interface AssignmentData {
  assignment:MicrosoftGraph.EducationAssignment;
  studentSubmissionDateStatus:string;
  currentPage:number;
  subjectName:string;
}

export interface AssignmentDataItemProps{
  itemData:AssignmentData;
  teamName:string;
  subject:string;
  course:string;
}

export interface IMyAssignmentsWebPartProps {
  pagingValue:number;
  hideOverDue:boolean;
  showArchivedTeams:boolean;
}



