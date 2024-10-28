import * as React from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './MQUser.module.scss';
import { IUser } from "../../model/Iuser";
import { IUserPredecessor } from './IUserProps';

export interface IUserPredecessorlocal
{
  MQPredecessorUser : IUser[];
  showPredecessor:boolean;
}

export default class MQUserPredecessors extends React.Component<IUserPredecessor,IUserPredecessorlocal, {}> 
{

 constructor(props: IUserPredecessor)
  {
    super(props);
    this.state={
        MQPredecessorUser :props.MQPredecessorData,
        showPredecessor:false 
    };
  }
  private toggle_Predecessor()
  {
    let currentstate = this.state.showPredecessor;
    this.setState({showPredecessor: !currentstate});
  }
  public render(): React.ReactElement<IUserPredecessor> {
    let icon =  this.state.showPredecessor ? <Icon iconName='chevronDown'/>:<Icon iconName='chevronUp'/>
      return (
        <div>
            <div className={styles.Mqpredecessors} ><span >Predecessors</span><button className="btn_predecessor" onClick={() => this.toggle_Predecessor()}>{icon}</button></div>
            <div style={{display: this.state.showPredecessor ? 'block' : 'none' }}>
            {this.props.MQPredecessorData.map((prdscr) => (
                            <>
                            <span className={styles.MqpredecessorprimaryText}>{prdscr.name}</span>
                            <span  title={prdscr.projectRole} className={styles.MqpredecessorSecondaryText}>{prdscr.projectRole} Since [ {new Date(prdscr.dateTo).toLocaleDateString()} ]</span>
                            <span className={styles.MqpredecessorSecondaryText}></span>
                            </>
            ))}

            </div>
        </div>
      );
    }
}