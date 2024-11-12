import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMqGraphConsumerProps {
  _Context:WebPartContext
}

export interface IMqGraphUserProps{
businessPhones: string,
displayName: string,
givenName: string,
id: string,
jobTitle: string,
mail: string,
mobilePhone: string,
preferredLanguage: string,
surname: string,
userPrincipalName:string 
}

export interface IUserPresence {
  activity?: "String",
  availability?:"String",
  id?: "String (identifier)",
  statusMessage?:{"@odata.type": "#microsoft.graph.presenceStatusMessage"}
}