/**
 * Validates that when Forward To is set to "Yes", at least one email address is selected in the multiselect PCF control
 * This function should be called on the form's OnSave event
 * 
 * @param {Object} executionContext - The execution context from Dynamics 365
 */
function validateForwardToEmails(executionContext) {
    try {
        // Get the form context
        var formContext = executionContext.getFormContext();
        
        // Get the Forward To field value (tec_forwardto)
        var forwardToField = formContext.getAttribute("tec_forwardto");
        var forwardToValue = forwardToField ? forwardToField.getValue() : null;
        
        // Get the PCF multiselect control value (tec_multiselectusers)
        var multiselectField = formContext.getAttribute("tec_multiselectusers");
        var selectedEmails = multiselectField ? multiselectField.getValue() : null;
        
        console.log("Forward To Value:", forwardToValue);
        console.log("Forward To Value Type:", typeof forwardToValue);
        console.log("Selected Emails:", selectedEmails);
        
        // Check if Forward To is set to "Yes" (900000000)
        // Note: Default value is "Select" which should not trigger validation
        // Possible values: null (Select), 900000000 (Yes), 900000001 (No)
        if (forwardToValue === 900000000) {
            // Check if no emails are selected or the field is empty
            if (!selectedEmails || selectedEmails.trim() === "") {
                // Prevent save and show error message
                executionContext.getEventArgs().preventDefault();
                
                // Show notification to user
                formContext.ui.setFormNotification(
                    "Please select at least one email address in the user selection field before saving when Forward To is set to Yes.",
                    "ERROR",
                    "forwardToEmailValidation"
                );
                
                // Optional: Focus on the PCF control
                var multiselectControl = formContext.getControl("tec_multiselectusers");
                if (multiselectControl) {
                    multiselectControl.setFocus();
                }
                
                return false;
            } else {
                // Valid selection - clear any existing notifications
                formContext.ui.clearFormNotification("forwardToEmailValidation");
                return true;
            }
        } else {
            // Forward To is not "Yes" (could be "Select", "No", null, etc.) - clear any existing notifications and allow save
            formContext.ui.clearFormNotification("forwardToEmailValidation");
            return true;
        }
        
    } catch (error) {
        console.error("Error in validateForwardToEmails:", error);
        
        // Show error to user
        var formContext = executionContext.getFormContext();
        formContext.ui.setFormNotification(
            "An error occurred while validating the forward to emails. Please try again.",
            "ERROR",
            "validationError"
        );
        
        // Prevent save on error
        executionContext.getEventArgs().preventDefault();
        return false;
    }
}

/**
 * Alternative function that can be used for real-time validation when Forward To field changes
 * This function should be called on the tec_forwardto field's OnChange event
 * 
 * @param {Object} executionContext - The execution context from Dynamics 365
 */
function onForwardToChange(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        
        // Get the Forward To field value
        var forwardToField = formContext.getAttribute("tec_forwardto");
        var forwardToValue = forwardToField ? forwardToField.getValue() : null;
        
        // If Forward To is set to "Yes", show a message encouraging user to select emails
        // Only show this message when explicitly set to "Yes", not for default "Select" value
        if (forwardToValue === 900000000) {
            var multiselectField = formContext.getAttribute("tec_multiselectusers");
            var selectedEmails = multiselectField ? multiselectField.getValue() : null;
            
            if (!selectedEmails || selectedEmails.trim() === "") {
                formContext.ui.setFormNotification(
                    "Please select email addresses for forwarding since Forward To is set to Yes.",
                    "INFO",
                    "forwardToReminder"
                );
            }
        } else {
            // Clear the reminder notification if Forward To is not "Yes"
            formContext.ui.clearFormNotification("forwardToReminder");
        }
        
    } catch (error) {
        console.error("Error in onForwardToChange:", error);
    }
}

/**
 * Function to validate when the multiselect field changes
 * This function should be called on the tec_multiselectusers field's OnChange event
 * 
 * @param {Object} executionContext - The execution context from Dynamics 365
 */
function onMultiselectChange(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        
        // Get current values
        var forwardToField = formContext.getAttribute("tec_forwardto");
        var forwardToValue = forwardToField ? forwardToField.getValue() : null;
        
        var multiselectField = formContext.getAttribute("tec_multiselectusers");
        var selectedEmails = multiselectField ? multiselectField.getValue() : null;
        
        // If Forward To is "Yes" and emails are now selected, clear any validation messages
        if (forwardToValue === 900000000 && selectedEmails && selectedEmails.trim() !== "") {
            formContext.ui.clearFormNotification("forwardToEmailValidation");
            formContext.ui.clearFormNotification("forwardToReminder");
        }
        
    } catch (error) {
        console.error("Error in onMultiselectChange:", error);
    }
}

/**
 * Utility function to check if the PCF control has valid email addresses
 * 
 * @param {string} emailString - Semicolon-separated email addresses
 * @returns {boolean} - True if at least one valid email exists
 */
function hasValidEmails(emailString) {
    if (!emailString || emailString.trim() === "") {
        return false;
    }
    
    // Split by semicolon and check if we have at least one non-empty email
    var emails = emailString.split(';');
    for (var i = 0; i < emails.length; i++) {
        if (emails[i].trim() !== "") {
            return true;
        }
    }
    
    return false;
}