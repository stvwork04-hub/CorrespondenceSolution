import * as React from 'react'; 
import { Label, DefaultButton } from '@fluentui/react'; 

// Props interface expects a parameter object for the first event
export interface IButtonEventProps {
    Id?: string;
    Name?: string;
    text?: string;
    textColour?: string;
    color?: string;
    tooltip?: string;
    onButtonClick?: (params: { Id: string; Name: string }) => void;
}

// Functional component
export const ButtonEvent: React.FunctionComponent<IButtonEventProps> = ({
    Id = "",
    Name = "",
    text = "Click me",
    textColour = "#ffffff",
    color = "#0078d4",
    tooltip = "Click to execute custom action",
    onButtonClick
}) => {
    const [isHovered, setIsHovered] = React.useState(false);
    const hoverColor = "#005a9e"; // Change this to your desired hover color
    return (
        <div>
            <DefaultButton
                onClick={() => {
                    if (onButtonClick) {
                        onButtonClick({ Id, Name });
                    }
                }}
                style={{
                    marginRight: "8px",
                    color: textColour,
                    backgroundColor: isHovered ? hoverColor : color,
                    transition: "background-color 0.2s"
                }}
                title={tooltip}
                onMouseEnter={() => setIsHovered(true)}
                onMouseLeave={() => setIsHovered(false)}
            >
                {text}
            </DefaultButton>
        </div>
    );
};