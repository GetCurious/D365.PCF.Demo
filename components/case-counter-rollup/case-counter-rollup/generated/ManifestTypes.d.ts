/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    caseStatus: ComponentFramework.PropertyTypes.EnumProperty<"0" | "1">;
    priorityCode: ComponentFramework.PropertyTypes.EnumProperty<"3" | "2" | "1">;
    caseDataset: ComponentFramework.PropertyTypes.DataSet;
}
export interface IOutputs {
}
