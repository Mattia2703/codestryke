import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from '../Findme.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DefaultButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { EProfileProperties, IUserProfilPlus} from '../models/IUserModel';
import * as strings from 'FindmeWebPartStrings';
import "@pnp/sp/profiles";
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { EditMultiStrings } from './EditMultiStrings';
import { useState } from 'react';
export interface ITempValues{
    availability: string;
    officeSeat: string;
    aboutMe: string;

}

export interface IViewFktEditProfileProps {
    webPartContext: WebPartContext;
    user: IUserProfilPlus;
    setView: () => void;
    setValues: (newValues: ITempValues) => void;
    setProfileProperty: (email: string, propertyName: string, properyValue: string, isMultiValue: boolean) => void;
    //handleAction:(action:EMenu,dataToSolve:string|undefined) => void;
}

const iconStyle = { fontSize: 25, marginRight: 25 };

const DateTime = () => <Icon style={iconStyle} iconName="DateTime" />;
const POI = () => <Icon style={iconStyle} iconName="POI" />;



const buttonStyles: Partial<IButtonStyles> = {
    root: {
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

const textFieldStyles: Partial<ITextFieldStyles> = {
    root: {
        
        
    },
    fieldGroup:{
        border:"0.5px solid #80808082"
    },
    wrapper:{
        //border:"0.5px solid #80808082"
    } 
}



export const ViewFktEditProfile: React.FunctionComponent<IViewFktEditProfileProps> = (props) => {


    const [availability, setAvail] = useState(props.user.Availability);
    const [officeSeat, setSeat] = useState(props.user.OfficeSeat);
    const [aboutMe, setAbout] = useState(props.user.AboutMe);
    const [skills, setSkills] = useState(props.user.Skills === "-" ? "" : props.user.Skills);
    const [projects, setProjects] = useState(props.user.Projects === "-" ? "" : props.user.Projects);
    const [interestes, setInterests] = useState(props.user.Interests === "-" ? "" : props.user.Interests);





    const setViewAndValue = () => {

        const newValues: ITempValues = {
            availability: availability,
            officeSeat: officeSeat,
            aboutMe: aboutMe,
        }

        props.setValues(newValues);


        props.setView()

    }



    return (
        <div className={styles.detailsContainer}>
            <div className={styles.editButtonContainer}>
                <DefaultButton
                    styles={buttonStyles}
                    iconProps={{ iconName: 'Save' }}
                    text={strings.DP_SaveProfile}
                    onClick={() => setViewAndValue()}
                />
            </div>
            <div className={styles.detailsComponentContainers}>
                <div className={styles.detailsTitle}>
                    {strings.DP_AvailabilityAndWorkplace}
                </div>
                <div className={styles.workPlaceContainer}>
                    <div className={styles.infoProfilCard}>
                        <div><DateTime /></div>
                        <TextField
                        className={styles.editBackground}
                        value = {availability}
                        multiline rows={3}
                        resizable={false}
                        styles={textFieldStyles}
                        onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => setAvail(newValue)}/>
                    </div>
                    <div className={styles.infoProfilCard}>
                        <div><POI /></div>
                        <TextField
                        className={styles.editBackground}
                        value = {officeSeat}
                        multiline rows={3}
                        resizable={false}
                        styles={textFieldStyles}
                        onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => setSeat(newValue)}
                        
                        />
                    </div>
                </div>
            </div>
            <div className={styles.detailsComponentContainers}>
                <div className={styles.detailsTitle}>
                    {strings.DP_TasksAndResponsibility}
                </div>
                <div className={styles.aboutMeContainer}>
                <TextField
                        className={styles.editBackground}
                        value = {aboutMe}
                        multiline rows={5}
                        resizable={false}
                        styles={textFieldStyles}
                        onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => setAbout(newValue)}/>
                </div>
            </div>
            <div className={styles.detailsComponentContainers}>
                <div className={styles.workPlaceContainer}>
                    <div className={styles.infoProfilCard}>
                        <div className={styles.detailsTitle}>{strings.DP_Qualification}</div>
                        <EditMultiStrings
                        webPartContext={props.webPartContext}
                        user={props.user}
                        setProfileProperty={props.setProfileProperty}
                        content={skills}
                        valueName={EProfileProperties.Skills}
                        setContent={setSkills}
                        textFieldStyles={textFieldStyles}/>
                    </div>
                    <div className={styles.infoProfilCard}>
                        <div className={styles.detailsTitle}>{strings.DP_Projects}</div>
                        <EditMultiStrings
                        webPartContext={props.webPartContext}
                        user={props.user}
                        setProfileProperty={props.setProfileProperty}
                        content={projects}
                        valueName={EProfileProperties.Projects}
                        setContent={setProjects}
                        textFieldStyles={textFieldStyles}
                        />
                    </div>
                </div>
            </div>
            <div className={styles.detailsComponentContainers}>
                <div className={styles.detailsTitle}>{strings.DP_Deputies}</div>
            </div>
            <div className={styles.detailsComponentContainers}>
                <div className={styles.detailsTitle}>{strings.DP_Interests}</div>
                <EditMultiStrings
                        webPartContext={props.webPartContext}
                        user={props.user}
                        setProfileProperty={props.setProfileProperty}
                        content={interestes}
                        valueName={EProfileProperties.Interests}
                        setContent={setInterests}
                        textFieldStyles={textFieldStyles}/>
            </div>
        </div>)


}
