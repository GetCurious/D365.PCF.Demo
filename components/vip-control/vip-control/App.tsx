import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    Button,
    Text,
    tokens,
} from "@fluentui/react-components";
import {PeopleStarRegular, PeopleTeamRegular} from "@fluentui/react-icons";
import {type IInputs} from "./generated/ManifestTypes";
import {useState} from "react";

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
                    <VIPComponent {...context.parameters} />
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function VIPComponent(props: IInputs) {
    const Xrm = React.useContext(XrmContext);

    const fontColor = tokens.colorStrokeFocus2; // black
    const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
    const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

    const [isVIP, setIsVIP] = useState(props.isVIP.raw)

    // onMount life-cycle
    React.useEffect(() => {
        // const result = Xrm.WebApi.retrieveRecord("contact", "55e7eb6f-3029-ef11-840a-000d3ac92c59");
        // Xrm.Navigation.openConfirmDialog({ text: "Hello" });
    }, []);


    return (
        <div style={{display: "flex", gap: "0.3rem"}}>
            <Button
                onClick={() => setIsVIP(true)}
                appearance={isVIP ? "primary" : "outline"}
                icon={<PeopleStarRegular fontSize={30} color={isVIP ? lightColor : primaryColor} />}
                shape="rounded">
                <Text as="h2">
                    <Text size={400} weight="bold" style={{color: (isVIP ? lightColor : primaryColor)}}>
                        VIP
                    </Text>
                    <Text size={100} block={true}>
                        &gt; 100,000.00 Revenue
                    </Text>
                </Text>
            </Button>
            <Button
                onClick={() => setIsVIP(false)}
                appearance={!isVIP ? "primary" : "outline"}
                icon={<PeopleTeamRegular fontSize={30} color={!isVIP ? lightColor : primaryColor} />}
                shape="rounded">
                <Text as="h2">
                    <Text size={400} weight="bold" style={{color: (!isVIP ? lightColor : primaryColor)}}>
                        Normal
                    </Text>
                    <Text size={100} block={true} style={{color: (!isVIP ? lightColor : fontColor)}}>
                        &lt; 100,000.00 Revenue
                    </Text>
                </Text>
            </Button>
        </div>
    );
}
