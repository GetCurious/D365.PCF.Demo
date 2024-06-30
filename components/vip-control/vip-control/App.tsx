import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    CompoundButton,
    tokens,
} from "@fluentui/react-components";
import {PeopleStarRegular, PeopleTeamRegular} from "@fluentui/react-icons";
import {type IInputs, IOutputs} from "./generated/ManifestTypes";

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
    rerender: (outputs: IOutputs) => void;
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
                    <VIPComponent params={context.parameters} rerender={this.props.rerender} />
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function VIPComponent(props: { params: IInputs, rerender: (outputs: IOutputs) => void }) {
    const Xrm = React.useContext(XrmContext);
    const [state, dispatch] = React.useReducer(reducer, {isVIP: props.params.isVIP.raw});

    const fontColor = tokens.colorStrokeFocus2; // black
    const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
    const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

    // onMount life-cycle
    React.useEffect(() => {
        // const result = Xrm.WebApi.retrieveRecord("contact", "55e7eb6f-3029-ef11-840a-000d3ac92c59");
        // Xrm.Navigation.openConfirmDialog({ text: "Hello" });
    }, []);


    // trigger update
    React.useEffect(() => {
        props.rerender(state);
    }, [state])


    return (
        <div style={{display: "flex", gap: "0.3rem"}}>
            <CompoundButton
                onClick={() => dispatch(Action.IS_VIP)}
                appearance={state.isVIP ? "primary" : "outline"}
                icon={<PeopleStarRegular />}
                secondaryContent={"> 100,000.00 Revenue"}
                shape="rounded">
                VIP
            </CompoundButton>
            <CompoundButton
                onClick={() => dispatch(Action.IS_NORMAL)}
                appearance={!state.isVIP ? "primary" : "outline"}
                icon={<PeopleTeamRegular />}
                secondaryContent={"< 100,000.00 Revenue"}
                shape="rounded">
                Normal
            </CompoundButton>
        </div>
    );
}


// events
enum Action {
    IS_NORMAL = 0,
    IS_VIP = 1,
}

function reducer(state: IOutputs, action: Action): IOutputs {
    switch (action) {
        case Action.IS_VIP:
            return {...state, isVIP: true};
        case  Action.IS_NORMAL:
            return {...state, isVIP: false};
        default:
            return state;
    }
}