import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI  } from "@pnp/sp";
import { GraphFI} from "@pnp/graph";
import "@pnp/graph/users";
import { getSP, getSPGraph } from "../helper/pnpjsConfig";
import {  IServiceResultDeputies } from "../MainPropsInterfacesAndEnums";
import "@pnp/sp/profiles";
import { EProfileProperties, IDeputiesProfil} from "../models/IUserModel";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { IUserDataFromGraphAADMin } from "./ServiceForProfile";
import HelperForProfile from "./HelperForProfile";


export default class ServiceForDeputies {

    private _webPartContext:WebPartContext;
    
    private _workingLog:string;
    private _hasError:boolean;
    private _sp:SPFI;
    private _graph:GraphFI

    constructor(webPartContext:WebPartContext){
        this._webPartContext = webPartContext;
        
        this._workingLog = "";
        this._hasError = false;
        this._sp = getSP(webPartContext);
        this._graph = getSPGraph();
    }


    /* public getProfilePic(upnMail: string){
       
        let pictureUrl = "/_vti_bin/DelveApi.ashx/people/profileimage?userId="+upnMail;
        pictureUrl += "&size=M";
        return(pictureUrl);
    }*/


    public async getUserByUPN(upnMail:string=undefined):Promise<IServiceResultDeputies>{
        let result:IServiceResultDeputies = undefined;
        console.log("load user im Service");
        result = { 
            isSuccess:false,            
            log:"",
            userObj:undefined
            
        } 
        try {
            let fullProfil:any = undefined;
            if(upnMail === undefined){
                result.log += "Lade myProperties\n";
                fullProfil = await this._sp.profiles.myProperties();
                //GET Azure AD
            } else {
                result.log += "Lade PropertiesFor "+upnMail+"\n";
                //loginName wird benötigt
                if(upnMail === undefined || upnMail.indexOf("@") <= -1){
                    throw "missing upnMail\n";
                }
                //let userToGetOrAddAndGet = await this._sp.web.ensureUser(upnMail);
                //GAST: "i:0#.f|membership|extern.rso_hb-munich.com#ext#@hbm2climb99.onmicrosoft.com
                //bei Externen muss das #ext# enthalten sein
                
                const loginName = "i:0#.f|membership|"+upnMail;

                fullProfil = await this._sp.profiles.getPropertiesFor(loginName);
                //gleichzeitig immer noch eine Azure Graph Anfrage (für bspw. restliche Felder wie Company)

            }
            const propsOfProfile:{[key:string]: string} = {};
            if(fullProfil.UserProfileProperties !== undefined){
                result.log += "Hole die Detail-Props\n ";
                fullProfil.UserProfileProperties.map((val:{Key:string,Value:string})=>{  
                         propsOfProfile[val.Key] = val.Value;  
                }); 
            }
            
            const userDataFromGraphAAD:IUserDataFromGraphAADMin = await this.getUserDataFromGraph(upnMail);

            result.userObj = this.prepareUserObjORM(fullProfil,propsOfProfile, userDataFromGraphAAD);
            if(result.userObj!==undefined){
                result.isSuccess = true;
            }    
         
        } catch(e)
        {
          result.log += "Fehler via try-catch abgefangen - ServiceForProfile\n" ;
          result.log += e
          console.error("Fail load profil");
        }



        return result;


    }
    private async getUserDataFromGraph(usermail:string):Promise<IUserDataFromGraphAADMin>{

        let graphResult;
        if(this._webPartContext.pageContext.user.email === usermail || usermail === undefined){
            graphResult = await this._graph.me.select("CompanyName")();

       }else {
        ///BUG
        ////https://github.com/pnp/pnpjs/issues/2238
           let usermailSave = usermail.replace("#","%23");
           usermailSave = usermailSave.replace("#","%23");
           //if(usermail.toLocaleLowerCase().indexOf("#ext#")>=0){
              /* let client = await this._webPartContext.msGraphClientFactory.getClient("3");
              graphResult = await client.api('/users/'+usermail)              
              .select('streetAddress, city, state, postalCode, country, CompanyName')              
              .get();*/
              //%23
           //} else {
                //https://graph.microsoft.com/v1.0/users/5cdab694-ac72-4812-81c3-412a893a536f
                
           //}
           ////
           graphResult = await this._graph.users.getById(usermailSave).select("CompanyName")();
       } 
       //bitte eine nullpointer Exception abfangen
       const userDataFromGraph: IUserDataFromGraphAADMin = {
        
         CompanyName: graphResult.companyName
       }

       return(userDataFromGraph);
    }


    private prepareUserObjORM(profile:any,profileProperties:any, userDataFromGraph?: IUserDataFromGraphAADMin):IDeputiesProfil{

        //Profile Picture
        //alle Felder bitte abfangen auf undefined / null!!!



        console.log(profileProperties);
        

        const userObj:IDeputiesProfil = {
            DisplayName:profile.DisplayName ? profile.DisplayName : "-",
            DepartmentShort:"ToDo",
            FirstName: profileProperties[EProfileProperties.FirstName] ? profileProperties[EProfileProperties.FirstName] : "-",
            LastName: profileProperties[EProfileProperties.LastName] ? profileProperties[EProfileProperties.LastName] : "-",
            aadUpnMail: profile[EProfileProperties.Email] ? profile[EProfileProperties.Email] : "-",
            Saeule: userDataFromGraph.CompanyName ? userDataFromGraph.CompanyName : "-",
            JobTitle: profileProperties[EProfileProperties.Title] ? profileProperties[EProfileProperties.Title] : "-",
            Department: profileProperties[EProfileProperties.Department] ? profileProperties[EProfileProperties.Department] : "-",
            isLoading:false,
            PictureUrls:HelperForProfile.getProfilPicturesByMail(profile[EProfileProperties.Email]),

            
        }



        return userObj;
    }



    public static getBlankDeputies(mail:string):IDeputiesProfil{
        
        return {
            DisplayName:"",
            DepartmentShort:"",
            FirstName: "",
            LastName: "",
            aadUpnMail: mail,
            Saeule: "",
            JobTitle: "",
            Department: "",
            PictureUrls:HelperForProfile.getProfilPicturesByMail(mail),
            isLoading:true
            
        }
    }

}