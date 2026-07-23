import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDataUploadProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}


// export interface IDataUploadProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
//   currentSPContext: any;
//   context?:any;
//   id:number;
// }
 
 