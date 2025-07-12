import * as React from 'react';
import styles from './SpFxAjax.module.scss';
import type { ISpFxAjaxProps } from './ISpFxAjaxProps';

export default class SpFxAjax extends React.Component<ISpFxAjaxProps> {
  public render(): React.ReactElement<ISpFxAjaxProps> {
    const {
      SPContext
    } = this.props;

    const _handleRead = (): void => {
      // Read data from ProjectStatus list and show in console.
      const siteabsulr = SPContext.pageContext.web.absoluteUrl
      const Listtitle = 'ProjectStatus'
      const requrl = siteabsulr + "/_api/web/lists/getbytitle('" + Listtitle + "')/items"

      fetch (requrl,{
        method: "GET",
        headers: {
                  'Accept': 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': ''
        }
      })
      .then(async response =>{
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        console.log('List items:', data.d.results);
      }) 
      alert("Read Operation complete");
    }

    const _handleUpdate = () : void => {
      const siteabsulr = SPContext.pageContext.web.absoluteUrl
      const Listtitle = 'ProjectStatus'
      const requrl = siteabsulr + "/_api/web/lists/getbytitle('" + Listtitle + "')/items(1)"

      const updatedItem = {
        __metadata: { type: 'SP.Data.ProjectStatusListItem' },
        Title: "Custom Title during Demo"
      };

      fetch (requrl,{
        method: "MERGE",
        headers: {
                  'Accept': 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': '',
                  'X-RequestDigest': this.props.SPContext.pageContext.legacyPageContext.formDigestValue,
                  'IF-MATCH': '*'
        },
        body: JSON.stringify(updatedItem)
      })
      .then(async response =>{
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        //const data = await response.json();
        alert("Read Operation updated");
      }) 
      alert("Read Operation complete");
    }

    return (
      <section>
        <div className={styles.welcome}>
          <h3> SPFx Ajax Demo</h3>
          <input type='text' id="ttxid"></input>
          <button onClick={_handleRead}> Read Data</button>
          <button onClick={_handleUpdate}>Update Data </button>
          {console.log(SPContext.pageContext.user.displayName)}
        </div>
      </section>
    );
  }
}
