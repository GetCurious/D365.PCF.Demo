import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    CompoundButton,
    Text,
    tokens,
} from "@fluentui/react-components";
import {MoneyCalculatorRegular} from "@fluentui/react-icons";
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
                    <OpportunityRevenue {...context.parameters} />
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function OpportunityRevenue(inputs: IInputs) {
    const Xrm = React.useContext(XrmContext);

    // const fontColor = tokens.colorStrokeFocus2; // black
    const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
    // const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

    enum OpportunityStatus {
        Open = 0,
        Won = 1,
    }

    const opportunityStatus = Number(inputs.opportunityStatus.raw || 0);
    const contactId = Xrm.entityId;

    const [totalRevenue, setTotalRevenue] = React.useState(0);

    // onMount life-cycle
    React.useEffect(() => {
        if (contactId === null) return;

        // for (const [id, entity] of opportunities) {
        //     let revenueValue = entity.getValue("estimatedvalue");
        //     if (revenueValue == null)
        //         Xrm.Navigation.openErrorDialog({message: "Dataset is missing logical column 'estimatedvalue'."});
        //     totalRevenue += Number(revenueValue);
        // }

        (async () => {
            const query = `?$select=estimatedvalue&$filter=(statecode eq ${opportunityStatus} and _parentcontactid_value eq '${contactId}')`;

            try {
                const {entities} = await Xrm.WebApi.retrieveMultipleRecords("opportunity", query);
                for (const entity of entities) {
                    setTotalRevenue(totalRevenue + entity.estimatedvalue);
                }
            } catch (err) {
                console.error(err);
            }
        })()
    }, [opportunityStatus, contactId]);

    return (
        <CompoundButton
            appearance={"outline"}
            secondaryContent={`${OpportunityStatus[opportunityStatus]} Revenue`}
            icon={<MoneyCalculatorRegular />}
            size="medium"
            shape="rounded">
            <Text style={{color: primaryColor}} size={400} weight={"bold"}>
                {totalRevenue.toLocaleString()}
            </Text>
        </CompoundButton>
    );
}
