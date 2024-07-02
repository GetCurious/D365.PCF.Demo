import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    Text,
    Spinner,
    Caption1Strong
} from "@fluentui/react-components";
import {
    GroupedVerticalBarChart,
    type IGroupedVerticalBarChartData,
} from '@fluentui/react-charting';
import {type IInputs, IOutputs} from "./generated/ManifestTypes";
import {StackShim} from "@fluentui/react-migration-v8-v9";


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
                    <PulzBarChart params={context.parameters} rerender={this.props.rerender}/>
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function PulzBarChart(props: { params: IInputs, rerender: (outputs: IOutputs) => void }) {
    const Xrm = React.useContext(XrmContext);
    const wonColor = "rgb(85, 216, 254)";
    const lostColor = "rgb(163, 160, 251)";

    const dataset = props.params.dataset;
    const columns = dataset.columns.map(c => c.name);
    const records = Object.values(dataset.records);
    const viewName = dataset.getTitle();

    if (!columns.includes("actualclosedate") || !columns.includes("estimatedvalue") || !columns.includes("statecode")) {
        console.warn(`one or more Field was not included in the view, consider adding it into ${viewName}.`);

        (dataset as any).addColumn("actualclosedate");
        (dataset as any).addColumn("estimatedvalue");
        (dataset as any).addColumn("statecode");
        dataset.refresh();

        return <Spinner size={"huge"} />
    }
    
    type Month = string
    type Revenue = {
        won: number;
        lost: number;
    };

    // aggregate
    const monthAggregate = new Map<Month, Revenue>();
    records.forEach(r => {
        const actualCloseDate = new Date(r.getValue("actualclosedate") as string);
        const revenue = r.getValue("estimatedvalue") as number;
        const status = r.getFormattedValue("statecode");
        const monthName = actualCloseDate.toLocaleString('default', {month: 'short'});
        if (!monthAggregate.has(monthName)) {
            monthAggregate.set(monthName, {won: 0, lost: 0})
        }

        let month = monthAggregate.get(monthName) as Revenue;
        if (status.toLowerCase() === 'won')
            month.won += revenue;
        else
            month.lost += revenue;
    });

    // build
    const data: IGroupedVerticalBarChartData[] = [];
    monthAggregate.forEach((point, month) => {
        const items: IGroupedVerticalBarChartData = {
            name: month,
            series: [
                {
                    key: 'lost',
                    data: point.lost,
                    xAxisCalloutData: 'Lost',
                    color: lostColor,
                    legend: 'Lost',
                },
                {
                    key: 'won',
                    data: point.won,
                    xAxisCalloutData: 'Won',
                    color: wonColor,
                    legend: 'Won',
                },
            ],
        };
        data.push(items)
    })

    return (
        <StackShim horizontalAlign={"start"}>
            <Caption1Strong style={{paddingLeft: "1rem"}}>
                Overall Sales
            </Caption1Strong>
            <GroupedVerticalBarChart
                chartTitle="Overall Sales"
                culture={window.navigator.language}
                data={data}
                legendProps={{
                    shape: "circle",
                    overflowProps: {
                        vertical: true
                    },
                }}
                barwidth={"auto"}
                yAxisTickCount={5}
                xAxisOuterPadding={0.8}
                secondaryYScaleOptions={{yMinValue: 20_000, yMaxValue: 100_000}}
            />
        </StackShim>

    );
}
