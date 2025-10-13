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
    private forwardTo: string = 'NO';

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
        
        // Initialize forwardTo from the input property
        this.forwardTo = context.parameters.forwardTo.raw || "NO";
        
        // Render the React component
        this.renderComponent();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        console.log('PCF Index - updateView called');
        
        this.context = context;
        
        // Update selectedEmails if the input value has changed
        const newEmailValue = context.parameters.selectedEmails.raw || "";
        const newForwardToValue = context.parameters.forwardTo.raw || "NO";
        
        console.log('PCF Index - updateView - Current field value:', newEmailValue);
        console.log('PCF Index - updateView - Current selectedEmails:', this.selectedEmails);
        console.log('PCF Index - updateView - Current forwardTo field value:', newForwardToValue);
        console.log('PCF Index - updateView - Current forwardTo:', this.forwardTo);
        
        if (newEmailValue !== this.selectedEmails) {
            console.log('PCF Index - updateView - Email field value changed, updating selectedEmails');
            this.selectedEmails = newEmailValue;
        }
        
        if (newForwardToValue !== this.forwardTo) {
            console.log('PCF Index - updateView - ForwardTo field value changed, updating forwardTo');
            this.forwardTo = newForwardToValue;
        }
        
        this.renderComponent();
    }

    /**
     * Handles email selection changes
     */
    private onEmailsChanged = (emails: string): void => {
        console.log('PCF Index - onEmailsChanged called with emails:', emails);
        console.log('PCF Index - Previous selectedEmails value:', this.selectedEmails);
        console.log('PCF Index - Previous forwardTo value:', this.forwardTo);
        
        this.selectedEmails = emails;
        
        // Update forwardTo based on whether emails are selected
        const hasEmails = emails && emails.trim().length > 0;
        this.forwardTo = hasEmails ? 'YES' : 'NO';
        
        console.log('PCF Index - Updated selectedEmails to:', this.selectedEmails);
        console.log('PCF Index - Updated forwardTo to:', this.forwardTo);
        console.log('PCF Index - Has emails:', hasEmails);
        
        // Add a small delay to ensure state is updated
        setTimeout(() => {
            this.notifyOutputChanged();
            console.log('PCF Index - notifyOutputChanged called');
            console.log('PCF Index - Final selectedEmails:', this.selectedEmails);
            console.log('PCF Index - Final forwardTo:', this.forwardTo);
        }, 100);
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
        console.log('PCF Index - getOutputs called');
        console.log('PCF Index - Current selectedEmails value:', this.selectedEmails);
        
        // Determine forwardTo value based on whether emails are selected
        const hasEmails = this.selectedEmails && this.selectedEmails.trim().length > 0;
        const forwardToValue = hasEmails ? 'YES' : 'NO';
        
        console.log('PCF Index - Has emails:', hasEmails);
        console.log('PCF Index - ForwardTo value:', forwardToValue);
        console.log('PCF Index - Returning output object:', { 
            selectedEmails: this.selectedEmails, 
            forwardTo: forwardToValue 
        });
        
        return {
            selectedEmails: this.selectedEmails,
            forwardTo: forwardToValue
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
