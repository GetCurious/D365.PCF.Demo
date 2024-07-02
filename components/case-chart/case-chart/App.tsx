import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    CompoundButton,
    tokens,
    Text,
} from "@fluentui/react-components";
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
    const [state, setState] = React.useState<IOutputs>({isVIP: props.params.isVIP.raw});

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
                onClick={() => setState({isVIP: true})}
                appearance={state.isVIP ? "primary" : "outline"}
                icon={<PeopleStarRegular color={state.isVIP ? lightColor : fontColor} />}
                secondaryContent={"> 100,000.00 Revenue"}
                shape="rounded"
                size={"small"}>
                <Text style={{color: state.isVIP ? lightColor : primaryColor}} size={400} weight={"semibold"}>
                    VIP
                </Text>
            </CompoundButton>
            <CompoundButton
                onClick={() => setState({isVIP: false})}
                appearance={!state.isVIP ? "primary" : "outline"}
                icon={<PeopleTeamRegular color={state.isVIP ? fontColor : lightColor} />}
                secondaryContent={"< 100,000.00 Revenue"}
                shape="rounded"
                size={"small"}>
                <Text style={{color: !state.isVIP ? lightColor : primaryColor}} size={400} weight={"bold"}>
                    Normal
                </Text>
            </CompoundButton>
        </div>
    );
}
