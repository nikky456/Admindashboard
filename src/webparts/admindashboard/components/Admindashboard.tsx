import * as React from "react";
import type { IAdmindashboardProps } from "./IAdmindashboardProps";
import Admin from "./Admin";
import { Web } from "sp-pnp-js";
import TrainerDashboard from "./Trainer/TrainerDashboard";

export default class Admindashboard extends React.Component<IAdmindashboardProps, { userType: string | null }> {
  constructor(props: IAdmindashboardProps) {
    super(props);
    this.state = {
      userType: null,
    };
  }

  componentDidMount() {
    this.fetchUserType();
  }

  private fetchUserType = async () => {
    try {
      const web = new Web("https://smalsusinfolabs.sharepoint.com/sites/F4S");

     
      const res: { AdminType: string; Admin: { Title: string; Id: number } }[] = await web.lists
        .getByTitle("AdminUsers")
        .items.select("AdminType", "Admin/Title", "Admin/Id")
        .expand("Admin")
        .get();

      const currentUserName = this.props.context?._legacyPageContext.userDisplayName; 
      let userType: string | null = null;

      
      for (const item of res) {
        if (item.Admin?.Title === currentUserName) {
          userType = item.AdminType;
          break; 
        }
      }

      this.setState({ userType });
    } catch (err) {
      console.error("Error fetching user type:", err);
    }
  };

  public render(): React.ReactElement<IAdmindashboardProps> {
    const { userType } = this.state;
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div>
        {userType === "Admin" && <Admin />}
        {userType === "Trainer" && (
          
          <TrainerDashboard
            description={description}
            isDarkTheme={isDarkTheme}
            environmentMessage={environmentMessage}
            hasTeamsContext={hasTeamsContext}
            userDisplayName={userDisplayName}
          />
        )}
        {!userType && <div>No dashboard available for this user.</div>}
      </div>
    );
  }
}
