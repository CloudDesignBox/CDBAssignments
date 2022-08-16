import * as React from 'react';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { AssignmentDataItemProps } from './IMyAssignmentsProps';
import styles from './MyAssignments.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

const stackTokens: Partial<IStackTokens> = { childrenGap: 20 };


export default class AssignmentItemDivV2 extends React.Component<AssignmentDataItemProps, {}> {

  
  constructor(props){
    super(props);
    this.openDetailsInDialog = this.openDetailsInDialog.bind(this);
  }

  private cleanDueDate:string="";

  private renderFriendlyDateFormat(datestring:string):string{
    let dateString:Date = new Date(datestring);
    const days = [
      'Sun',
      'Mon',
      'Tue',
      'Wed',
      'Thu',
      'Fri',
      'Sat'
    ];
    const months = [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'May',
      'Jun',
      'Jul',
      'Aug',
      'Sep',
      'Oct',
      'Nov',
      'Dec'
    ];
    let hours:number = dateString.getHours();
    let hr:string = hours < 10 ? '0' + hours.toString() : hours.toString();

    let minutes:number = dateString.getMinutes();
    let min:string = (minutes < 10) ? '0' + minutes.toString() : minutes.toString();
    let newTimeString:string = hr + ':' + min;
    let day:string = dateString.getDate().toString();
    const monthName = months[dateString.getMonth()];
    const dayName = days[dateString.getDay()];
    return `${dayName} ${day} ${monthName} at ${newTimeString}`;
  }

  private openDetailsInDialog(){
    window.open(`https://teams.microsoft.com/l/entity/66aeee93-507d-479a-a3ef-8f494af43945/classroom?context=%7B%22subEntityId%22%3A%22%7B%5C%22version%5C%22%3A%5C%221.0%5C%22,%5C%22config%5C%22%3A%7B%5C%22classes%5C%22%3A%5B%7B%5C%22id%5C%22%3A%5C%22${this.props.itemData.assignment.classId}%5C%22,%5C%22displayName%5C%22%3A%5C%22AssignmentsCalendar%5C%22,%5C%22assignmentIds%5C%22%3A%5B%5C%22${this.props.itemData.assignment.id}%5C%22%5D%7D%5D%7D,%5C%22action%5C%22%3A%5C%22navigate%5C%22,%5C%22view%5C%22%3A%5C%22assignment-viewer%5C%22%7D%22,%22channelId%22%3Anull%7D`,"_blank");
  }

private truncateString(words:string, num:number){
  if (words.length > num) {
    let subStr:string = words.substring(0, num);
    return subStr + "...";
  } else {
      return words;
  }
}

/* eslint-disable react/jsx-no-bind */
public render(): React.ReactElement<AssignmentDataItemProps> {
    //validate data
    let displayName:string = "No title";
    let instructions:string = "";
    let statusColour:string = styles.black;
    let subjectInitials:string = "H";
    let subjectName:string = "";
    let teamTitle:string="";
    let alerticon = <span />;
    if(this.props.itemData){
        if(this.props.itemData.assignment.dueDateTime){
            this.cleanDueDate = this.renderFriendlyDateFormat(this.props.itemData.assignment.dueDateTime);
        }
        if(this.props.itemData.assignment.displayName){
            displayName = this.truncateString(this.props.itemData.assignment.displayName,20);
        }
        if(this.props.itemData.assignment.instructions){
            if(this.props.itemData.assignment.instructions.content){
            instructions = this.props.itemData.assignment.instructions.content;
            }
        }
        if(this.props.itemData.studentSubmissionDateStatus == "overdue"){
          statusColour=styles.red;
          alerticon=<span><Icon iconName="Clock" className={styles.alerticon} />&nbsp;</span>;
        }
    }

    if(this.props.course){
      subjectInitials=this.props.course.substring(0,1);
      subjectName = this.truncateString(this.props.course,8)+" - ";
    }else if(this.props.subject){
      subjectInitials=this.props.subject.substring(0,1);
      subjectName = this.truncateString(this.props.subject,8)+" - ";
    }else if(this.props.subject){
      subjectInitials=this.props.subject.substring(0,1);
    }

    if(this.props.teamName){
      teamTitle = this.truncateString(this.props.teamName,35);
    }
    // isArchived not returned?
    let classesForItem:string=`${styles.assignmentOuterBlock} cdbassignmentpage${this.props.itemData.currentPage.toString()}`;
  return (
    <div className={classesForItem}>
      <div onClick={this.openDetailsInDialog} className={styles.assignmentBlock}>
        <div className={styles.assignmentIcon}>{subjectInitials}</div>
        <div className={styles.assignmentDesc}>
          <b>{subjectName}{displayName}</b><br />
           {alerticon}<span className={statusColour}>Due {this.cleanDueDate}</span><br />
           {teamTitle}
        </div>
      </div>
      <div className={styles.clear} />
    </div>
  );
}
}