import * as React from "react";
import { IUsercollection } from "../../model/Iusercollection";
import { IUser } from "../../model/Iuser";
import nodestyles from "./Orgchartnode.module.scss";
import styles from "../MqOrgchartHierarchy.module.scss";
import MQUser from "../user/MQUser";

import { DocumentCard, DocumentCardActions, DocumentCardDetails } from '@fluentui/react';

interface Ichilduser {
  userdata: IUser;
}

export default class OrgChartNode extends React.Component<Ichilduser, {}> {

  constructor(props: any) {
    super(props);
    this.state = {
      User: props.userdata
    };
  }
  private _onClick(action: string, ev:React.SyntheticEvent<HTMLElement>):void {
    alert(`You clicked the ${action} action`);
    window.open(
        `https://teams.microsoft.com/l/chat/0/0?users=nilesh.ubhe@marquardt.com&message=Hi `,
        "_blank"
      );
    ev.stopPropagation();
    ev.preventDefault();
}
  public render(): React.ReactElement<IUsercollection> {

    return (
      <li>
        <div className={styles.node + " " + (this.props.userdata.children && this.props.userdata.children.length > 0 ? "" : styles["no-child"]) }>
          
          <DocumentCard>
            <DocumentCardDetails>
                <MQUser MQUserData={this.props.userdata}></MQUser>
            </DocumentCardDetails>
            <DocumentCardActions className={nodestyles.DocumentActionsroot} actions={
              [
                {
                  iconProps: { iconName: 'Chat' },
                  onClick: this._onClick.bind(this, 'Chat'),
                  ariaLabel: 'Chat'
                },
                {
                  iconProps: { iconName: 'Mail' },
                  onClick: this._onClick.bind(this, 'Mail'),
                  ariaLabel: 'Mail'
                }/*,
                {
                  iconProps: { iconName: 'Phone' },
                  onClick: this._onClick.bind(this, 'Phone'),
                  ariaLabel: 'Phone'
                },*/
              ]
            } >
            </DocumentCardActions>
          </DocumentCard>

        </div>
        {this.props.userdata.children  && (
        <ul className={styles["sub-list"]}>
          {this.props.userdata.children.map((child) => (
            <OrgChartNode userdata={child} />
          ))}
  </ul>
)}

      </li>
    );
  }
}