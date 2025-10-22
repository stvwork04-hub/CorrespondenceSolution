import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ButtonEvent, IButtonEventProps } from "./components/ButtonEvent";

export class TECCustomEventButton implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement | null = null;
    constructor() {
        // intentionally empty 
    }
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, 
           state: ComponentFramework.Dictionary, 
           container: HTMLDivElement): void {
        // store the container provided by the platform
        this.container = container;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if (!this.container) return;
        // Read the PCF input properties
        const id = context.parameters.Id.raw || "";
        const name = context.parameters.Name.raw || "";
        const buttonText = context.parameters.buttonText.raw || "Click me";
        const buttonTextColour = context.parameters.buttonTextColour.raw || "#ffffff";
        const buttonColor = context.parameters.buttonColor.raw || "#0078d4";
        const buttonTooltip = context.parameters.buttonTooltip.raw || "Click to execute custom action";

        // Prepare props for your React component
        const props: IButtonEventProps & { text: string; textColour: string; color: string; tooltip: string } = {
            Id: id,
            Name: name,
            onButtonClick: () => { context.events.onButtonClick({ Id: id, Name: name }); },
            text: buttonText,
            textColour: buttonTextColour,
            color: buttonColor,
            tooltip: buttonTooltip
        };

        // Render the React component into the container
        ReactDOM.render(React.createElement(ButtonEvent, props), this.container);
    }
    
    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        // Clean up the React component when PCF is removed 
        if (this.container) {
            ReactDOM.unmountComponentAtNode(this.container);
        }
    }
}