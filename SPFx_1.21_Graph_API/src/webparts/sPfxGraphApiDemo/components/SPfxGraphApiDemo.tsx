import * as React from 'react';
import type { ISPfxGraphApiDemoProps } from './ISPfxGraphApiDemoProps';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react';
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";

export interface IState
{
  users: any[];
}

export default class SPfxGraphApiDemo extends React.Component<ISPfxGraphApiDemoProps,IState> {

  constructor(props:ISPfxGraphApiDemoProps )
  {
    super(props);
    this.state = {
      users:[]
    };
  }

  protected async onInit(): Promise<void> {
    return Promise.resolve();
  }

  public async componentDidMount(): Promise<void> {
    const { webpartcontext } = this.props;
    const graph = graphfi().using(SPFx(webpartcontext));
    const users = await graph.users();
    this.setState({ users });
  }

  public render(): React.ReactElement<ISPfxGraphApiDemoProps> {
    const columns : IColumn[] = [
      { key: 'displayName', name: 'Display Name', fieldName: 'displayName', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'mail', name: 'Email', fieldName: 'mail', minWidth: 150, maxWidth: 300, isResizable: true },
      { key: 'jobTitle', name: 'Job Title', fieldName: 'jobTitle', minWidth: 100, maxWidth: 200, isResizable: true },
    ];
    return (
      <section >
          <DetailsList
            items={this.state.users || []}
            columns={columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
      </section>
    );
  }
}
