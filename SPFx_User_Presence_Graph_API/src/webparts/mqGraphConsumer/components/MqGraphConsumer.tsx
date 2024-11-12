import * as React from 'react';
import styles from './MqGraphConsumer.module.scss';
import type { IMqGraphConsumerProps } from './IMqGraphConsumerProps';
import { GraphFI, graphfi, SPFx as graphSPFx } from '@pnp/graph'
import "@pnp/graph/users";
import { IMqGraphUserProps } from './IMqGraphConsumerProps';
import { IUserPresence } from './IMqGraphConsumerProps';
import "@pnp/graph/cloud-communications";


export interface IMqUserState {
  Userid?: string;
  activity?:string;
  availability?:string;
}

export default class MqGraphConsumer extends React.Component<IMqGraphConsumerProps, IMqUserState> {

  private _graph : GraphFI ;
  
  public async getPresence(userId: string): Promise<IUserPresence> {
    
    return new Promise<IUserPresence>(async (resolve, reject) => {
      try {
        let MQUserpresenseresponse:any = await this._graph.users.getById(userId).presence();
        console.log(MQUserpresenseresponse)
        resolve(MQUserpresenseresponse);
      }
      catch (error) {
        reject(error); // Reject the promise with the error
      }
    });
  }

  public async getUserId(useremail?: string): Promise<IMqGraphUserProps> {
    return new Promise<IMqGraphUserProps>(async (resolve, reject) => {
      try {
        if (useremail)
        {
        let userid: IMqGraphUserProps = await this._graph.users.getById(useremail)();
        resolve(userid);
        }
        else
        {
          const emptyUserProps: IMqGraphUserProps = {
            businessPhones: '',
            displayName: '',
            givenName: '',
            id: '',
            jobTitle: '',
            mail: '',
            mobilePhone: '',
            preferredLanguage: '',
            surname: '',
            userPrincipalName: '',
        };
          resolve(emptyUserProps);
        }
      }
      catch (error) {
        reject(error); // Reject the promise with the error
      }
    });
  }

  constructor(props: IMqGraphConsumerProps) {
    super(props);
    this.state = {
      Userid: "",
      activity:"",
      availability:""
    };
  }

  
  public componentDidMount(): void {
    this._graph = graphfi().using(graphSPFx(this.props._Context));
    this.getuserafterpromise();
  }

  private getuserafterpromise(): void {
    // Get the current user details
      this.getUserId("MeganB@0nthj.onmicrosoft.com")
      .then((_userid: IMqGraphUserProps) => {
        this.getPresence(_userid.id)  
        .then((userpresense: IUserPresence) => {
              this.setState({activity : userpresense.activity, availability: userpresense.availability, Userid: userpresense.id });
        });
      });
  }


  public render(): React.ReactElement<IMqGraphConsumerProps> {

    return (
    <section className={`${styles.mqGraphConsumer}`}>
    <h2> User id - {this.state?.Userid} </h2>
    <h3> Activity - {this.state?.activity} </h3>
    <h3> Availability - {this.state?.availability} </h3>
    </section>
    )
   
  }
}
