import * as React from 'react';
import IUserProps from './IUserProps';
import styles from './MQUser.module.scss';
import MQUserPredecessors from './MQUserPredecessor';
//import MQUserPresense from './MQUserpresense';

export default class MQUser extends React.Component<IUserProps, {}> {

  constructor(props: any) {
    super(props);
    this.state = {
      MQUser: props.User
    };
  }

  public render(): React.ReactElement<IUserProps> {

    return (
      <div>
        <table className={styles.cardTable}>
          <tr>
            <td>
          <div className="ms-Image ms-Persona-image image-233" >
            <img className={styles.PersonaIamge}
              src={this.props.MQUserData.userimage}
              alt=""></img>
          </div>
          </td>
          </tr>
          <tr>
            <td>
          <div dir="auto" className="MquserprimaryText">
            <span className={styles.MquserprimaryText}>{this.props.MQUserData.name}</span>
            <span title={this.props.MQUserData.projectRole} className={styles.MquserSecondaryText}>{this.props.MQUserData.projectRole} </span>
            <span className={styles.MquserSecondaryText}>Since{new Date(this.props.MQUserData.dateTo).toLocaleDateString()}</span>
            
          </div>
          {this.props.MQUserData.predecessor && this.props.MQUserData.predecessor.length > 0 && <MQUserPredecessors MQPredecessorData={this.props.MQUserData.predecessor} showPredecessor={false}></MQUserPredecessors>}
          </td>
          </tr>
        </table>
      </div>
    );
  }
}