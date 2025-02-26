import * as React from "react";
import type { ITrainerDashboardProps } from "./ITrainerDashboardProps";
// import Table from "./Table";
import Trainer_Dashboard from "./Trainer_Dashboard";

export default class TrainerDashboard extends React.Component<ITrainerDashboardProps> {
  public render(): React.ReactElement<ITrainerDashboardProps> {
    return (
      <Trainer_Dashboard />
      // <Table />
    );
  }
}
