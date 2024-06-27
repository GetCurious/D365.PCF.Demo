import * as React from "react";
import {
  FluentProvider,
  webLightTheme,
  PartialTheme,
  Badge,
  Text,
  tokens,
} from "@fluentui/react-components";
import { MoneyCalculatorRegular } from "@fluentui/react-icons";
import { type IInputs } from "./generated/ManifestTypes";

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
function OpportunityRevenue(props: IInputs) {
  const Xrm = React.useContext(XrmContext);

  const fontColor = tokens.colorStrokeFocus2; // black
  const primaryColor = tokens.colorBrandForegroundInverted; // rgb(71, 158, 245)
  const lightColor = tokens.colorNeutralBackground2; // rgb(250, 250, 250)

  enum OpportunityStatus {
    Open = 0,
    Won = 1,
  }

  const opportunities = Object.entries(props.opportunityDataset.records);
  const opportunityStatus = props.opportunityStatus.raw;
  const contactId = Xrm.entityId;
  let totalRevenue = 0;

  // onMount life-cycle
  React.useEffect(() => {
    for (const [id, entity] of opportunities) {
      let revenueValue = entity.getValue("estimatedvalue");
      if (revenueValue == null)
        Xrm.Navigation.openErrorDialog({message: "Dataset is missing logical column 'estimatedvalue'."});
      totalRevenue += Number(revenueValue);
    }
    
    // const query = `?$select=estimatedvalue&$filter=(statecode eq '${opportunityStatus}' and _parentcontactid_value eq '${contactId}')`;
    // Xrm.WebApi.retrieveMultipleRecords("opportunity", query)
    //   .then(({ entities }) => {
    //     for (const entity of entities) {
    //       totalRevenue += entity.estimatedvalue;
    //     }
    //   })
    //   .catch((err) => {
    //     console.error(err);
    //     Xrm.Navigation.openErrorDialog({ details: err });
    //   });
    
  }, [opportunities]);

  return (
    <Badge
      icon={<MoneyCalculatorRegular fontSize={30} color={fontColor} />}
      size="extra-large"
      shape="rounded"
      style={{
        padding: "1.8rem 0.5rem",
        backgroundColor: lightColor,
        boxShadow: tokens.shadow2,
      }}>
      <Text as="h2">
        <Text size={400} weight="bold" style={{ color: primaryColor }}>
          {totalRevenue}
        </Text>
        <Text size={100} block={true} style={{ color: fontColor }}>
          {OpportunityStatus[opportunityStatus]} Revenue
        </Text>
      </Text>
    </Badge>
  );
}
