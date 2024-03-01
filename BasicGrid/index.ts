import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IGridProps, DataGrid } from "./BasicGrid";
import * as React from "react";

export class BasicGrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;

        const props: IGridProps = {
            gridHeight: "350px",
            dataset: this.context.parameters.dataset,
            navigate: this.navigate.bind(this),
            setSelected: this.setSelected.bind(this),
        };
        return React.createElement(
            DataGrid, props
        );
    }

    public navigate(options: ComponentFramework.NavigationApi.EntityFormOptions) {
        this.context.navigation.openForm(options);
    }

    public setSelected = (ids: string[]): void => {
        this.context.parameters.dataset.setSelectedRecordIds(ids);
    };

    public getOutputs(): IOutputs {
        return { };
    }

    public destroy(): void {
    }
}
