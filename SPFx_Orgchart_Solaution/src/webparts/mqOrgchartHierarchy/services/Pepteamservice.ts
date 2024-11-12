
import { IUser } from "../model/Iuser";
import { Ipepteamservice } from "../model/Ipepteamservice";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { GraphFI, graphfi } from "@pnp/graph";
import "@pnp/graph/users";
import { IUserPresence } from "../components/user/IUserPresense";

export class Pepteamservice implements Ipepteamservice {

  private _sp: SPFI;
  private _graph : GraphFI ;

  constructor(private context: WebPartContext) {
    this._sp = spfi().using(SPFx(this.context));
  }

  public async getUserId(useremail?: string): Promise<string> {
    this._graph = graphfi().using(SPFx(this.context));
    return new Promise<string>(async (resolve, reject) => {
      try {
        if (useremail)
        {
        let userid: string = await this._graph.users.getById(useremail)();
        resolve(userid);
        }
        else
        {
          resolve("empty")
        }
      }
      catch (error) {
        reject(error); // Reject the promise with the error
      }
    });
} 
  
public async getPresence(userId: string): Promise<IUserPresence> {
  let MQUserpresenseInstance : IUserPresence;

  return new Promise<IUserPresence>(async (resolve, reject) => {
    try {
      const MQUserpresenseresponse = await  this._graph.users.getById(userId).presence();
      MQUserpresenseInstance = {
        Activity :  MQUserpresenseresponse.activity?.toString(),
        Availability: MQUserpresenseresponse.availability?.toString()
      }
      resolve(MQUserpresenseInstance);
    }
    catch (error) {
      reject(error); // Reject the promise with the error
    }
  });
}


  public async getAllTeamUsers(siteurl: string,): Promise<IUser[]> {

    return new Promise(async (resolve, reject) => {
      try {
        let Datarry: IUser[] = [];
        let data: any[] = await this._sp.web.lists.getByTitle("Team").items.orderBy("SuperiorRole", true).orderBy("ProjectRole", true).orderBy("ResponsibleTo", false)();;

        if (data) {
          for (let i = 0; i < data.length; i++) {
            let Username = '';
            let Useremail = '';
            let Userimage = '';
            if (data[i].MySiteUserId) {
              const userData = await this.getUserName(data[i].MySiteUserId);
              console.log(userData)
              Username = userData[0].user.Title;
              Useremail = userData[0].user.Email
              Userimage = userData[0].profileImageUrl;
            }
            Datarry.push({
              name: Username,
              email:Useremail,
              projectRole: data[i].ProjectRole,
              superiorRole: data[i].SuperiorRole,
              dateFrom: data[i].ResponsibleFrom,
              dateTo: data[i].ResponsibleTo,
              userimage: Userimage
            });
          }
        }
        resolve(Datarry); // Resolve the promise with the Datarry array
      } catch (error) {
        reject(error); // Reject the promise with the error
      }
    });
  }

  private async getUserName(MySite_x0020_UserId: number) {
    const userArray = [];
    const user = await this._sp.web.getUserById(MySite_x0020_UserId).select('Title', 'Email')();
    const profileImageUrl = `${window.location.protocol}//${window.location.hostname}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(user.Email)}`;
    //const userid = await this.getUserId(user.Email);
    userArray.push({ user, profileImageUrl});
    return userArray;
  }

  private findMatchingInaciveRoles(Projectdata: IUser[], InactiveProjectRole: string): IUser[] {
    const InactiveProjectRoles: IUser[] = [];
    Projectdata.forEach(project => {
      let projectdateto = new Date(project.dateTo);
      if (project.children && project.children.length > 0) {

        project.children = this.findMatchingInaciveRoles(project.children, InactiveProjectRole);
        if (project.projectRole == InactiveProjectRole && projectdateto < new Date()) {
          InactiveProjectRoles.push(project);
        }
      }
    });
    return InactiveProjectRoles;
  }

  private removeEmptyChildren(data: IUser[]): IUser[] {
    data.forEach(user => {
      if (user.children && user.children.length > 0) {
        user.children = this.removeEmptyChildren(user.children);
        if (user.children.length === 0) {
          delete user.children;
        }
      } else {
        delete user.children;
      }
    });
    return data;
  }

  public generateProjectCard(projects: any[]): IUser[] {
    const projectMap: { [key: string]: IUser } = {};
    const projectHierarchy: IUser[] = [];
    projects.forEach((project: any) => {
        const { projectRole, superiorRole, name, dateFrom, userimage, dateTo, email} = project;
        const Indexuser: IUser = {
          name,
          email,
          projectRole,
          superiorRole,
          dateFrom,
          dateTo ,
          userimage,
          children: [],
          predecessor:[]
        };

        if (projectMap[projectRole] && new Date(dateTo) < new Date())
          {
            projectMap[projectRole].predecessor?.push(Indexuser);
          }
        else if(new Date(dateTo) >= new Date())
         {
          projectMap[projectRole] = Indexuser;
        }
    });   

    Object.values(projectMap).forEach(user => {
        const { superiorRole } = user;

        if (superiorRole) {
            const superiorUser = projectMap[superiorRole];
            if (superiorUser) {
                superiorUser.children?.push(user);
            }
        } else {
            projectHierarchy.push(user);
        }
    });
    return this.removeEmptyChildren(projectHierarchy);
}

}
