import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ButtonEvent, IButtonEventProps } from "./components/ButtonEvent";

export class TECCustomEventButton implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement | null = null;
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this.container = container;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if (!this.container) return;
        const buttonText = context.parameters.buttonText.raw || "Click me";
        const buttonTextColour = context.parameters.buttonTextColour.raw || "#ffffff";
        const buttonColor = context.parameters.buttonColor.raw || "#0078d4";
        const buttonTooltip = context.parameters.buttonTooltip.raw || "Click to execute custom action";

        const props = {
            text: buttonText,
            textColour: buttonTextColour,
            color: buttonColor,
            tooltip: buttonTooltip
        };
        ReactDOM.render(React.createElement(ButtonEvent, props), this.container);
    }
    
    public getOutputs(): IOutputs { return {}; }

    public destroy(): void {
        if (this.container) {
            ReactDOM.unmountComponentAtNode(this.container);
        }
    }
}