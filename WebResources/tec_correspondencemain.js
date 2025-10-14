/**
 * Combined validation for "Forward To" and Multiselect PCF field.
 * Handles both OnSave and OnChange logic with clean error handling.
 */

/**
 * OnSave validation function for Forward To field requirements
 * @param {Object} executionContext - The execution context from Dynamics 365
 * @returns {boolean} - Returns true if validation passes, false otherwise
 */
function validateForwardToOnSave(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        
        // Get the Forward To field value
        var forwardToAttribute = formContext.getAttribute("tec_forwardto");
        var forwardToValue = forwardToAttribute ? forwardToAttribute.getValue() : null;
        
        // Get the Email address field value
        var emailAttribute = formContext.getAttribute("tec_multiselectusers");
        var emailValue = emailAttribute ? emailAttribute.getValue() : null;
        
        // Get controls for field-level notifications
        var forwardToControl = formContext.getControl("tec_forwardto");
        var emailControl = formContext.getControl("tec_multiselectusers");
        
        console.log("Forward To Value:", forwardToValue);
        console.log("Email Value:", emailValue);
        
        // Clear all existing notifications first (both field-level and form-level)
        clearForwardToValidationError(formContext);
        clearEmailValidationError(formContext);
        formContext.ui.clearFormNotification("VALIDATION_ERROR");
        
        // Validation Rule 1 & 2: Forward To should not be "Select" (default)
        // Assuming "Select" has value null, undefined, or 900000002
        if (forwardToValue === null || forwardToValue === undefined || forwardToValue === 900000002) {
            
            // Set field-level notification on Forward To control
            if (forwardToControl) {
                forwardToControl.setNotification(
                    "You need to select Yes or No.",
                    "FORWARD_TO_SELECT_ERROR"
                );
                forwardToControl.setFocus();
            }
            
            // Also set form-level notification as backup
            formContext.ui.setFormNotification(
                "You need to select Yes or No.",
                "ERROR",
                "FORWARD_TO_SELECT_ERROR"
            );
            
            // Prevent save
            executionContext.getEventArgs().preventDefault();
            return false;
        }
        
        // Validation Rule 3: If Forward To is "Yes" and Email is null/empty
        // Assuming "Yes" has value 900000000
        if (forwardToValue === 900000000) {
            if (isEmailFieldEmpty(emailValue)) {
                
                // Set field-level notification on Email control
                if (emailControl) {
                    emailControl.setNotification(
                        "Please select atleast one email address",
                        "EMAIL_REQUIRED_ERROR"
                    );
                    emailControl.setFocus();
                }
                
                // Also set form-level notification
                formContext.ui.setFormNotification(
                    "Please select atleast one email address",
                    "ERROR",
                    "EMAIL_REQUIRED_ERROR"
                );
                
                // Prevent save
                executionContext.getEventArgs().preventDefault();
                return false;
            }
        }
        
        // All validations passed - clear any remaining notifications
        clearForwardToValidationError(formContext);
        clearEmailValidationError(formContext);
        
        console.log("✅ Forward To validation passed");
        return true;
        
    } catch (error) {
        console.error("Error in validateForwardToOnSave:", error);
        
        // Show generic error message
        formContext.ui.setFormNotification(
            "A validation error occurred. Please try again.",
            "ERROR",
            "VALIDATION_ERROR"
        );
        
        // Prevent save on error
        executionContext.getEventArgs().preventDefault();
        return false;
    }
}

/**
 * REUSABLE FUNCTION: Clears Forward To validation errors
 * @param {Object} formContext - The form context
 */
function clearForwardToValidationError(formContext) {
    var forwardToControl = formContext.getControl("tec_forwardto");
    
    // Clear field-level notification
    if (forwardToControl) {
        forwardToControl.clearNotification("FORWARD_TO_SELECT_ERROR");
    }
    
    // Clear form-level notification
    formContext.ui.clearFormNotification("FORWARD_TO_SELECT_ERROR");
    
    console.log("Cleared Forward To validation error");
}

/**
 * REUSABLE FUNCTION: Clears Email validation errors and reminders
 * @param {Object} formContext - The form context
 */
function clearEmailValidationError(formContext) {
    var emailControl = formContext.getControl("tec_multiselectusers");
    
    // Clear field-level notifications
    if (emailControl) {
        emailControl.clearNotification("EMAIL_REQUIRED_ERROR");
        emailControl.clearNotification("EMAIL_REMINDER");
    }
    
    // Clear form-level notifications
    formContext.ui.clearFormNotification("EMAIL_REQUIRED_ERROR");
    formContext.ui.clearFormNotification("forwardToReminder");
    
    console.log("Cleared Email validation errors");
}

/**
 * Triggered when Forward To value changes.
 * Clears field-level and form-level notifications when valid selection is made.
 */
function onForwardToChange(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var forwardToAttr = formContext.getAttribute("tec_forwardto");
        var forwardToVal = forwardToAttr ? forwardToAttr.getValue() : null;
        
        var emailControl = formContext.getControl("tec_multiselectusers");

        console.log("Forward To changed to:", forwardToVal);

        // Clear Forward To validation errors if user selects anything other than "Select"
        if (forwardToVal !== null && forwardToVal !== undefined && forwardToVal !== 900000002) {
            console.log("Clearing Forward To notifications...");
            clearForwardToValidationError(formContext);
        }

        // Handle email validation when Yes is selected
        var multiAttr = formContext.getAttribute("tec_multiselectusers");
        var multiVal = multiAttr ? multiAttr.getValue() : null;

        if (forwardToVal === 900000000) { // "Yes"
            if (isEmailFieldEmpty(multiVal)) {
                // Show reminder on email field
                if (emailControl) {
                    emailControl.setNotification(
                        "Please select email addresses for forwarding",
                        "EMAIL_REMINDER"
                    );
                }
                formContext.ui.setFormNotification(
                    "Please select email addresses for forwarding since Forward To is set to Yes.",
                    "INFO",
                    "forwardToReminder"
                );
            } else {
                // Clear notifications if emails are already selected
                clearEmailValidationError(formContext);
            }
        } else {
            // Clear all email-related notifications if Forward To is not "Yes"
            clearEmailValidationError(formContext);
        }

    } catch (e) {
        console.error("Error in onForwardToChange:", e);
    }
}

/**
 * Triggered when Multiselect field changes.
 * Clears notifications once valid emails are selected.
 */
function onMultiselectChange(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var forwardToAttr = formContext.getAttribute("tec_forwardto");
        var forwardToVal = forwardToAttr ? forwardToAttr.getValue() : null;

        var multiAttr = formContext.getAttribute("tec_multiselectusers");
        var multiVal = multiAttr ? multiAttr.getValue() : null;

        console.log("Multiselect changed, value:", multiVal);

        if (forwardToVal === 900000000 && hasValidEmails(multiVal)) {
            console.log("Clearing email notifications...");
            clearEmailValidationError(formContext);
        }

    } catch (e) {
        console.error("Error in onMultiselectChange:", e);
    }
}

/**
 * Helper function to check if email field is empty or null
 * @param {string|array|null} emailValue - The email field value
 * @returns {boolean} - Returns true if empty/null, false if has valid emails
 */
function isEmailFieldEmpty(emailValue) {
    // Check for null or undefined
    if (!emailValue) {
        return true;
    }
    
    // If it's an array, check if it has elements
    if (Array.isArray(emailValue)) {
        return emailValue.length === 0;
    }
    
    // If it's a string, check if it's empty or just whitespace
    if (typeof emailValue === "string") {
        var trimmedValue = emailValue.trim();
        if (trimmedValue === "") {
            return true;
        }
        
        // If it's a semicolon-separated string, check if it has valid emails
        var emailArray = trimmedValue.split(";")
            .map(email => email.trim())
            .filter(email => email !== "");
        
        return emailArray.length === 0;
    }
    
    // For any other type, consider it empty
    return true;
}

/**
 * Utility: Works for both array or semicolon-separated email strings.
 */
function hasValidEmails(value) {
    if (!value) return false;

    if (Array.isArray(value)) return value.length > 0;

    if (typeof value === "string") {
        var arr = value.split(";").map(e => e.trim()).filter(e => e !== "");
        return arr.length > 0;
    }

    return false;
}