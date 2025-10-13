import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { SystemUserForwardComponent } from "./components";
import { SystemUser } from "./helpers";

export class MultiselectLookup implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;
    private selectedEmails: string = '';

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        
        // Initialize selectedEmails from the input property
        this.selectedEmails = context.parameters.selectedEmails.raw || "";
        
        // Render the React component
        this.renderComponent();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this.context = context;
        
        // Update selectedEmails if the input value has changed
        const newValue = context.parameters.selectedEmails.raw || "";
        if (newValue !== this.selectedEmails) {
            this.selectedEmails = newValue;
        }
        
        this.renderComponent();
    }

    /**
     * Handles email selection changes
     */
    private onEmailsChanged = (emails: string): void => {
        console.log('PCF Index - onEmailsChanged:', emails);
        this.selectedEmails = emails;
        this.notifyOutputChanged();
        console.log('PCF Index - notifyOutputChanged called');
    };

    /**
     * Renders the React component
     */
    private renderComponent(): void {
        const reactElement = React.createElement(SystemUserForwardComponent, {
            context: this.context,
            initialEmails: this.selectedEmails,
            onEmailsChanged: this.onEmailsChanged
        });

        ReactDOM.render(reactElement, this.container);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs
    {
        console.log('PCF Index - getOutputs called, returning:', this.selectedEmails);
        return {
            selectedEmails: this.selectedEmails
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        ReactDOM.unmountComponentAtNode(this.container);
    }

    /**
     * Gets the currently selected emails
     */
    public getSelectedEmails(): string {
        return this.selectedEmails;
    }

    /**
     * Sets the selected emails programmatically
     */
    public setSelectedEmails(emails: string): void {
        this.selectedEmails = emails;
        this.renderComponent();
        this.notifyOutputChanged();
    }
}
