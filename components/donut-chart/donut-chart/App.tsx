import * as React from "react";
import {
    FluentProvider,
    webLightTheme,
    PartialTheme,
    Text,
    Spinner,
    Caption1Strong
} from "@fluentui/react-components";
import {DonutChart, type IChartDataPoint} from '@fluentui/react-charting';
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

export class App extends React.PureComponent<IAppProps> {
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
                    <PulzDonutChart params={context.parameters} rerender={this.props.rerender}/>
                </XrmContext.Provider>
            </FluentProvider>
        );
    }
}

// === Actual Component =========================================
function PulzDonutChart(props: { params: IInputs, rerender: (outputs: IOutputs) => void }) {
    const Xrm = React.useContext(XrmContext);
    const openColor = "rgb(85, 216, 254)";
    const wonColor = "rgb(255, 131, 115)";
    const lostColor = "rgb(255, 218, 131)";
    const cancelledColor = "rgb(163, 160, 251)";
    const extraColor = "rgb(249, 192, 30)";
    const colors = [openColor, wonColor, lostColor, cancelledColor, extraColor];

    const targetField = props.params.targetField.raw?.trim() ?? "";
    const dataset = props.params.dataset;
    const columns = dataset.columns.map(c => c.name);
    const records = Object.values(dataset.records);
    const chartTitle = (dataset as any)?.entityDisplayCollectionName ?? dataset.getTargetEntityType();
    const viewName = dataset.getTitle();
    
    if (!targetField)
        return <Text>Target field is required</Text>

    if (!columns.includes(targetField)) {
        console.warn(`Field ${targetField} was not included in the view, consider adding it into ${viewName}.`);

        if (dataset.addColumn && targetField)
            dataset.addColumn(targetField);

        dataset.refresh()
        return <Spinner size={"huge"}/>
    }

    let optionSets = new Map<string, number>();

    console.log("Loaded columns: " + columns);
    records.forEach(r => {
        const formattedValue = r.getFormattedValue(targetField);
        if (!optionSets.get(formattedValue))
            optionSets.set(formattedValue, 0);

        optionSets.set(formattedValue, (optionSets.get(formattedValue) as number) + 1);
    });

    const points = Array.from(optionSets).map(([key, value], i) => ({
        legend: key,
        data: value,
        color: colors[i]
    }));

    return (
        <StackShim horizontalAlign={"end"}>
            <Caption1Strong style={{textTransform: "capitalize", paddingRight: "1rem"}}>
                {records.length} {chartTitle}
            </Caption1Strong>
            <DonutChart height={200} width={180} innerRadius={45}
                        culture={window.navigator.language}
                        data={{
                            chartTitle: chartTitle,
                            chartData: points,
                        }}
                        valueInsideDonut={""}
                        hideLabels={true}
                        hideLegend={false}
                        hideTooltip={false}
                        legendProps={{
                            shape: "circle",
                            overflowProps: {
                                vertical: true
                            },
                        }}
                        styles={{
                            root: {
                                flexDirection: "row",
                                columnGap: "0.5rem",
                            },
                            chart: {},
                            legendContainer: {}
                        }}>
            </DonutChart>
        </StackShim>

    );
}
