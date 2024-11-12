import * as React from "react";
import type { IMqOrgchartHierarchyProps } from "./IMqOrgchartHierarchyProps";
import OrgChartNode from "./Orgchartnode/orgchartnode";
import { IUser } from "../model/Iuser";
import styles from "./MqOrgchartHierarchy.module.scss";
import { Pepteamservice } from "../services/Pepteamservice";
import { TransformWrapper, TransformComponent, } from "react-zoom-pan-pinch";
import OrgchartToolBar from "./Toolbar/OrgchartToolBar";


export interface IMqOrgchartHierarchyState {
  userCollection: IUser[];
}

export default class MqOrgchartHierarchy extends React.Component<
  IMqOrgchartHierarchyProps,
  IMqOrgchartHierarchyState
> {
  private pepteamserviceinstance: Pepteamservice;
  private projectCards: IUser[];


  constructor(props: IMqOrgchartHierarchyProps) {
    super(props);
    this.state = {
      userCollection: [],
    };
    this.pepteamserviceinstance = new Pepteamservice(this.props.context);
    //this.obj =  this.pepteamserviceinstance.getAllTeamUsers("fdafd").then(response:)
  }

  public componentDidMount(): void {
    this.getuserafterpromise();
  }

  //Process orgchart data collection
  

  private getuserafterpromise(): void {
    // Get the current user details
    void this.pepteamserviceinstance
      .getAllTeamUsers("test")
      .then((users: IUser[]) => {
        this.projectCards = this.pepteamserviceinstance.generateProjectCard(users);
        console.log(this.projectCards);
        this.setState({ userCollection: this.projectCards });
      });
  }

  public render(): React.ReactElement<IMqOrgchartHierarchyProps> {
    return (
      <TransformWrapper
        initialScale={0.8}
        initialPositionX={100}
        initialPositionY={100}
        zoomIn={{ step: 0.2 }}
        zoomOut={{ step: 0.2 }}
      >
        {({ zoomIn, zoomOut, resetTransform }: any) => (
          <>
            <OrgchartToolBar
              zoomIn={zoomIn}
              zoomOut={zoomOut}
              resetTransform={resetTransform}
            />
            <TransformComponent>
              <div className={styles.orgchart} id="orgchart">
                <div id="toolbarcontainer"></div>

                <ul className={styles["prime-list"]}>
                  {this.state.userCollection.map((child) => (
                    <OrgChartNode userdata={child} />
                  ))}
                </ul>
              </div>
            </TransformComponent>
          </>
        )}
      </TransformWrapper>
    );
  }
}
