import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHacktheworldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext
}
export interface RolesandUsersProps {
  context:WebPartContext
}

