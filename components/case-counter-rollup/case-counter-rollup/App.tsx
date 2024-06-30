import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    CompoundButton,
    Text,
    tokens,
} from "@fluentui/react-components";
import {DocumentBriefcaseRegular, PeopleStarRegular} from "@fluentui/react-icons";
import {type IInputs} from "./generated/ManifestTypes";

// === Dependency Injection =====================================
export interface Xrm {
    Utility: ComponentFramework.Utility;
    WebApi: ComponentFramework.WebApi;
    Navigation: ComponentFramework.Navigation;
    userSettings: ComponentFramework.UserSettings;
    entityId: string;
    entityTypeName: string;
}

export const XrmContext = React.createContext({} as Xrm);

// === Theme Override ===========================================
const overrideTheme: PartialTheme = {
    ...webLightTheme,
    // Override here
    // colorBrandBackground: "rgb(71, 158, 245)",
};

// === Wrapper for setups (theme, context, etc) =================
export interface IAppProps {
    context: ComponentFramework.Context<IInputs>;
}

export class App extends React.Component<IAppProps> {
    public render(): React.ReactNode {
        const context = this.props.context;

        // setup global context
        const Xrm: Xrm = {
            Utility: context.utils,
            WebApi: context.webAPI,
            Navigation: context.navigation,
            userSettings: context.userSettings,
            // @ts-expect-error: https://powerusers.microsoft.com/t5/Power-Apps-Pro-Dev-ISV/How-to-access-the-underlying-entity-in-Standard-control/td-p/388921
            entityId: context.mode.contextInfo?.entityId,
            // @ts-expect-error: https://powerusers.microsoft.com/t5/Power-Apps-Pro-Dev-ISV/How-to-access-the-underlying-entity-in-Standard-control/td-p/388921
            entityTypeName: context.mode.contextInfo?.entityTypeName,
        };

        return (
            <FluentProvider theme={overrideTheme}>
                <XrmContext.Provider value={Xrm}>
                    <CaseCounterComponent {...context.parameters} />
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function CaseCounterComponent(props: IInputs) {
    const Xrm = React.useContext(XrmContext);

    // const fontColor = tokens.colorStrokeFocus2; // black
    // const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
    // const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

    enum CaseStatus {
        Open = 0,
        Resolved = 1,
    }

    enum PriorityCode {
        High = 1,
        Normal = 2,
        Low = 3,
    }

    // const cases = Object.entries(props.caseDataset.records);
    const caseStatus = Number(props.caseStatus.raw || 0);
    const priorityCode = Number(props.priorityCode.raw || 1);
    const contactId = Xrm.entityId;
    let [totalCase, setTotalCase] = React.useState(0);

    // onMount life-cycle
    React.useEffect(() => {
        if (contactId === null) return;

        // for (const [incidentId, entity] of cases) {
        //   totalCase += 1;
        //   
        // let priorityCodeValue = entity.getValue("prioritycode");
        // let stateCodeValue = entity.getValue("statecode");
        // let contactValue = entity.getValue("primarycontactid");
        // if (contactValue == null || stateCodeValue == null || priorityCodeValue == null) {
        //   Xrm.Navigation.openErrorDialog({message: "Dataset is missing logical column 'prioritycode'."});
        // }
        // }

        (async () => {
            const query = `?$select=incidentid&$filter=(statecode eq ${caseStatus} and prioritycode eq ${priorityCode} and _primarycontactid_value eq '${contactId}')`;

            try {
                let {entities} = await Xrm.WebApi.retrieveMultipleRecords("incident", query);
                for (const _ of entities) {
                    setTotalCase(totalCase + 1);
                }
            } catch (err) {
                console.error(err);
            }
        })()
    }, [caseStatus, priorityCode, contactId]);

    return (
        <CompoundButton
            appearance={"outline"}
            secondaryContent={`Priority ${CaseStatus[caseStatus]} Cases`}
            icon={<DocumentBriefcaseRegular />}
            size="large"
            shape="rounded">
            {totalCase} {PriorityCode[priorityCode]}
        </CompoundButton>
    );
}
