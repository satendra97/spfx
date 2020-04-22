import * as React from 'react';
import { IDemoreactProps } from './IDemoreactProps';
export default class Demoreact extends React.Component<IDemoreactProps, {}> {
    state: {
        ActionCategory: any[];
        associatedTasks: any[];
        project: any[];
        actions: any[];
        phase: any[];
        choose: {
            key: string;
            text: string;
        }[];
        Selectedproject: any;
        SelectedprojectId: any;
        Selectedactions: any;
        SelectedactionsId: any;
        SelectedassociatedTasks: any;
        SelectedassociatedTasksId: any;
        Selectedphase: any;
        SelectedphaseId: any;
        Selectedchoose: any;
        SelectedchooseId: any;
        SelectedUsers: any[];
        SelectedUserId: any;
        SelectedDate1: any;
        SelectedDate2: any;
        SelectedDate3: any;
    };
    render(): React.ReactElement<IDemoreactProps>;
    componentDidMount(): Promise<void>;
    private GetIPAddress;
    edit(): void;
    Submit(): void;
    private Cancel;
    private _getPeoplePickerItems;
    private handleChange1;
    private handleChange2;
    private handleChange3;
    private getProjects;
    private getactnitems;
    private getphases;
    private getTasks;
    private onChangeProject;
    private onChangeAction;
    private onChangeAssociatedTasks;
    private onChangePhase;
    private onChangechoice;
    private GetUserId;
    private GetUserDetails;
    private getid;
}
//# sourceMappingURL=Demoreact.d.ts.map