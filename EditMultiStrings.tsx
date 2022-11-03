import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from '../Findme.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DefaultButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { IUserProfilPlus } from '../models/IUserModel';
import { useState } from 'react';
import "@pnp/sp/profiles";

import { ITextFieldStyles, TextField } from '@fluentui/react/lib/TextField';
import { TooltipHost } from '@fluentui/react';
import {
   
    MessageBar,
    MessageBarType,
  
    MessageBarButton,
  
  } from '@fluentui/react';


export interface IEditMultiStringsProps {
    user: IUserProfilPlus;
    webPartContext: WebPartContext;
    content: string;
    valueName: string;
    textFieldStyles:Partial<ITextFieldStyles>;
    setContent: React.Dispatch<React.SetStateAction<string>>;
    setProfileProperty: (email: string, propertyName: string, properyValue: string, isMultiValue: boolean) => void;
    //handleAction:(action:EMenu,dataToSolve:string|undefined) => void;
}

const iconStyle = { fontSize: 12};

const Remove = () => <Icon style={iconStyle} iconName="Delete" />;






const buttonStyles: Partial<IButtonStyles> = {
    root: {
        minWidth: "20px",
        padding: "3px",
        border: "0px",
        backgroundColor: "#ffcc00",
    },
    label: {

    },
    rootHovered: {
        backgroundColor: "#d0a700",
    },
    rootPressed: {
        backgroundColor: "#d0a700",
    }
};

const backgroundColor: React.CSSProperties = {

    width: "100%",
}



export const EditMultiStrings: React.FunctionComponent<IEditMultiStringsProps> = (props) => {

    

    const [fieldContent, setField] = useState("");
    const [showAdded, setAdded] = useState(false);
    const [showDeleted, setDeleted] = useState(false);




    const deleteValue = (mapContent: string, removeValue: string) => {


        const splitContent = mapContent.split("|");

        let newString: string = ""; //create new String with deleted values

        splitContent.map((content, index) => {
            console.log(content, removeValue);
            if(content !== removeValue){
                newString += (content + "|")
            }
        });

        newString = newString.slice(0, -1);

        props.setProfileProperty(props.user.aadUpnMail, props.valueName, newString, true)

        props.setContent(newString)

        setAdded(false);
        setDeleted(true);
    }


    const addValue = (mapContent: string) => {

        if(fieldContent.length > 0){

            const newString: string = mapContent === "" ? fieldContent :  mapContent + "|" + fieldContent; //create new string


            props.setContent(newString); //Set new Content

            props.setProfileProperty(props.user.aadUpnMail, props.valueName, newString, true)

            setField(""); //Set Textfield to 0

            setDeleted(false);
            setAdded(true);
        }

        
        

        
    }


    const createObject = (mapContent: string) => {


        const splitContent = mapContent.split("|");

        const returnOBJ = splitContent.map((content, index) => {

            const obj = content !== "" ? 
                <div>
                    <div className={styles.singleValue}>
                        <div className={styles.mapTextOverflow}>
                            <TooltipHost content={content}>
                                {content}
                            </TooltipHost>
                        </div>
                        <div
                            onClick={() => deleteValue(mapContent, content)}
                            className={styles.removeIcon}>
                            <Remove />
                        </div>

                    </div>

                </div>:
                <div>
                <div className={styles.singleValue}>
                    <div className={styles.mapTextOverflow}>
                        <TooltipHost content={"No content"}>
                            {"No content"}
                        </TooltipHost>
                    </div>
                    <div
                        onClick={() => deleteValue(mapContent, content)}
                        className={styles.removeIcon}>
                        <Remove />
                    </div>

                </div>

            </div>;
            
            return (obj);
        });

        return (returnOBJ);

    }
    

    return (
        <div className={styles.multiValueEdit}>
            <div className={styles.multiValueControls}>
                <div style={backgroundColor}>
                    <TextField
                    value={fieldContent}
                    onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => setField(newValue)}
                     styles={props.textFieldStyles}
                     />
                </div>
                <DefaultButton
                    styles={buttonStyles}
                    iconProps={{ iconName: 'Add' }}
                    onClick={() => addValue(props.content)}
                />
            </div>
            <div className={styles.singleValuesContainer}>
                { showAdded &&
                  <MessageBar
                    actions={
                        <div>
                            <MessageBarButton
                            onClick={() => setAdded(false)}
                            >Ok</MessageBarButton>
                        </div>
                    }
                    messageBarType={MessageBarType.success}
                    isMultiline={false}
                >
                    Item Added
                </MessageBar>  
                }
                { showAdded || showDeleted &&
                  <MessageBar
                    actions={
                        <div>
                            <MessageBarButton
                            onClick={() => setDeleted(false)}
                            >Ok</MessageBarButton>
                        </div>
                    }
                    messageBarType={MessageBarType.success}
                    isMultiline={false}
                >
                    Item Deleted
                </MessageBar>  
                }
                
                {createObject(props.content)}
            </div>
        </div>)


}
