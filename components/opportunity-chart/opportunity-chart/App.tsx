import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    type BrandVariants,
    type Theme,
    CompoundButton,
    tokens,
    Text,
} from "@fluentui/react-components";
import {createV8Theme} from '@fluentui/react-migration-v8-v9';
import {ThemeContext_unstable as V9ThemeContext} from '@fluentui/react-shared-contexts';
import {DonutChart, type IChartProps, type IChartDataPoint} from '@fluentui/react-charting';
import {type IInputs, IOutputs} from "./generated/ManifestTypes";
import * as d3Color from 'd3-color';

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

    const fontColor = tokens.colorStrokeFocus2; // black
    const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
    const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

    let parentV9Theme = React.useContext(V9ThemeContext) as Theme;
    let v9Theme: Theme = parentV9Theme ? parentV9Theme : webLightTheme;
    let backgroundColor = d3Color.hsl(v9Theme.colorNeutralBackground1);
    let foregroundColor = d3Color.hsl(v9Theme.colorNeutralForeground1);
    const myV8Theme = createV8Theme(brandInvariant, v9Theme, backgroundColor.l < foregroundColor.l); // For dark theme background color is darker than foreground color

    const points: IChartDataPoint[] = [
        {legend: 'first', data: 20000, color: fontColor, xAxisCalloutData: '2020/04/30'},
        {
            legend: 'second',
            data: 39000,
            color: primaryColor,
            xAxisCalloutData: '2020/04/20',
        },
    ];
    const data: IChartProps = {
        chartTitle: 'Donut chart basic example',
        chartData: points,
    };

    return (
        <DonutChart
            culture={window.navigator.language}
            data={data}
            innerRadius={55}
            href={'https://developer.microsoft.com/en-us/'}
            legendsOverflowText={'overflow Items'}
            hideLegend={false}
            height={220}
            width={176}
            valueInsideDonut={39000}
        />
    );
}

const brandInvariant: BrandVariants = {
    10: '',
    20: '',
    30: '',
    40: '',
    50: '',
    60: '',
    70: '',
    80: '',
    90: '',
    100: '',
    110: '',
    120: '',
    130: '',
    140: '',
    150: '',
    160: '',
};