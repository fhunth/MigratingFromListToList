import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

//npm install @pnp/core --save
// npm i @pnp/sp
// import styles from './HelloWorldWebPart.module.scss';
// import { timer } from 'rxjs';
// import { take } from 'rxjs/operators';
// import * as $ from "jquery";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { spfi, SPFI, SPFx } from "@pnp/sp";



import {
  SPHttpClient,
  SPHttpClientResponse, 
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IItemUpdateResult } from '@pnp/sp/items';
// import { isNull } from 'lodash';

export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface SPListRecords {
  value: SPListItemRecord[];
}
export interface SPListItemRecord {
  ID: number;
  Title: string;
  SourceID:number;
  CC:number[];
  CCAUX:string;
  Fecharealdecierre:Date  ;
  fecharealdecierrestring:string  ;

}

export interface SPListFA01Destination {
  NameFA01:string;
  Description:string;
  Area:string;
  BudgetItem:string;
  ProjectNumber:string;
  AmountUSD:number;
  AmountLocal:number;
  Date:Date;
  Status:string;
  Region:string;
  Location:string;
  Project_x0020_Type:string;
  CASA_x0020_Type:string;
  PjctSponsorDisp:string;
  _x0020_DesiredDa:string;
  VoidControl:string;
  LeaseRequired:string;
  Budget_x0020_Item_x0020_Name:string;
  BudgetedAmmount:number;
  Budget_x0020_Next_x0020_Yr:number;
  Budget_x0020_In_x0020_2_x0020_Yr:string;
  Start_x0020_Date:Date;
  EndDate:Date;
  Owner:string;
  
  Void_x0020_Control:string;
  Budgeted:string;
  Company:string;
  Item:string;
  ProjectObjectives:string;
  CASA:string;
  CompanyNumber:string;  
  Country:string;
  FA01Description:string;
  Notices_x0020_Sent:string;
  NoticesSent:string;
  Pending_x0020_Approver:string;
  Purpose:string;
  QuotesJustification:string;
  Title:string;
  _x0020_Actual_x0020_Responsible:string;
  Actual_x0020_Responsible_x0020_Disp:string;
  Amount_x0020_USD:number;
  _x0020_Approval_x0020_Date:string;
  Approval1:string;
  Approval10:string;
  _x0020_Approval2:string;
  Approval3:string;
  Approval4:string;
  Approval:string;
  Approval7:string;
  Approval8:string;
  _x0020_Approval9:string;
  _x0020_Approver10Comments:string;
  Approver1Comments:string;
  Approver2Comments:string;
  Approver3Comments:string;
  _x0020_Approver4Comments:string;
  Approver5Comments:string;
  Approver6Comments:string;
  Approver7Comments:string;
  Approver8Comments:string;
  Approver9Comments:string;
  Aprvr10AprvlDate:string;
  Aprvr1AprvlDate:string;
  _x0020_Aprvr2AprvlDate:string;
  Aprvr3AprvlDate:string;
  Aprvr4AprvlDate:string;
  Aprvr5AprvlDate:string;
  Aprvr6AprvlDate:string;
  Aprvr7AprvlDate:string;
  Aprvr8AprvlDate:string;
  Aprvr9AprvlDate:string;
  BudgetedJustification:string;
  BudgetedPast:string;
  BudgetedPastJustification:string;
  Casa_x0020_Project:string;
  CASA_x0020_Ratio:string;
  CASA_x0020_Total:string;
  CASERatio:string;
  CASETotal:string;
  DFCIncluded:string;
  DFCJustificacion:string;
  _x0020_DocIncluded:string;
  DocJustification:string;
  Expansion_x0020_Ratio:string;
  ExpantionRatio:number;
  ExpantionTotal:number;
  ExpensesLocal:number;
  ExpensesUSD:number;
  GermanyMapping:string;
  GrandTotal:number;
  Id:number;
  ID2:string;
  Internal_x0020_Payback_x0020_Minimun:string;
  Is_x0020_CASA:string;
  Is_x0020_cheked_x0020_out_x0020_to_x0020_local:string;
 
  Lease_x0020_User:string;
  LeaseApproval:string;
  LeaseDisp:string;
  LeasePending:string;
  LSAprvlDate:string;
  LvBIncluded:string;
  LvBJustification:string;
  _x0020_Name2:string;
  Non_x0020_CASA_x0020_Total:string;
  ObsLeaseAuth:string;
  _x0020_Old_x0020_Form:string;
  OUBudget:string;
  Q_x0020_1_x0020_Budget:number;
  Q_x0020_2_x0020_Budget:number;
  _x0020_Q_x0020_3_x0020_Budget:number;
  Q_x0020_4_x0020_Budget:number;
  _x0020_Quotes:string;
  RecurringLocal:string;
  RecurringUSD:string;
  RejectionJustificatio:string;
  _x0020_Restart_x0020_Approval:string;
  RiskAssessmen:string;
  RiskAssessmentRequired:string;
  StartDate:string;
  State2:string;
  Sustenance_x0020_Total:number;
  SustenanceRatio:number;
  SustenanceTotal:string;
  TestLink:string;
  _x0020_Total:number;
  UsuariosCC:string;
  
}
//_x0020_

export interface SPListFA01Requests {
  value: SPListFA01Request[];
}
export interface SPListFA01Request {
  NameFA01:string;
Description:string;
Area:string;
BudgetItem:string;
ProjectNumber:string;
AmountUSD:number;
AmountLocal:number;
Date:string;
Status:string;
Region:string;
Location:string;
Project_x0020_Type:string;
CASA_x0020_Type:string;
PjctSponsorDisp:string;
DesiredDa:string;
VoidControl:string;
LeaseRequired:string;
Budget_x0020_Item_x0020_Name:string;
BudgetedAmmount:number;
Budget_x0020_Next_x0020_Yr:number;
Budget_x0020_In_x0020_2_x0020_Yr:string;
Start_x0020_Date:string;
EndDate:string;
Owner:string;

Void_x0020_Control:string;
Budgeted:string;
Company:string;
Item:string;
ProjectObjectives:string;
CASA:string;
CompanyNumber:string;

Country:string;
FA01Description:string;
Folder_x0020_Child_x0020_Count	:string;
Notices_x0020_Sent:string;
NoticesSent:string;
Pending_x0020_Approver:string;
Purpose:string;
QuotesJustification:string;
Title:string;
Actual_x0020_Responsible:string;
Actual_x0020_Responsible_x0020_Disp:string;
Amount_x0020_USD:number;
Approval_x0020_Date:string;
Approval1:string;
Approval10:string;
Approval2:string;
Approval3:string;
Approval4:string;
Approval:string;

Approval7:string;
Approval8:string;
Approval9:string;
Approver10Comments:string;
Approver1Comments:string;
Approver2Comments:string;
Approver3Comments:string;
Approver4Comments:string;
Approver5Comments:string;
Approver6Comments:string;
Approver7Comments:string;
Approver8Comments:string;
Approver9Comments:string;
Aprvr10AprvlDate:string;
Aprvr1AprvlDate:string;
Aprvr2AprvlDate:string;
Aprvr3AprvlDate:string;
Aprvr4AprvlDate:string;
Aprvr5AprvlDate:string;
Aprvr6AprvlDate:string;
Aprvr7AprvlDate:string;
Aprvr8AprvlDate:string;
Aprvr9AprvlDate:string;
BudgetedJustification:string;
BudgetedPast:string;
BudgetedPastJustification:string;
Casa_x0020_Project:string;
CASA_x0020_Ratio:string;
CASA_x0020_Total:string;
CASERatio:string;
CASETotal:string;
DFCIncluded:string;
DFCJustificacion:string;
DocIncluded:string;
DocJustification:string;
Expansion_x0020_Ratio:string;
ExpantionRatio:number;
ExpantionTotal:number;
ExpensesLocal:number;
ExpensesUSD:number;
GermanyMapping:string;
GrandTotal:number;
Id:number;
ID2:string;
Internal_x0020_Payback_x0020_Minimun:string;
Is_x0020_CASA:string;
Is_x0020_cheked_x0020_out_x0020_to_x0020_local:string;


Lease_x0020_User:string;
LeaseApproval:string;
LeaseDisp:string;
LeasePending:string;
LSAprvlDate:string;
LvBIncluded:string;
LvBJustification:string;
Name2:string;
Non_x0020_CASA_x0020_Total:string;
ObsLeaseAuth:string;
Old_x0020_Form:string;
OUBudget:string;
Q_x0020_1_x0020_Budget:number;
Q_x0020_2_x0020_Budget:number;
Q_x0020_3_x0020_Budget:number;
Q_x0020_4_x0020_Budget:number;
Quotes:string;
RecurringLocal:string;
RecurringUSD:string;
RejectionJustificatio:string;
Restart_x0020_Approval:string;
RiskAssessmen:string;
RiskAssessmentRequired:string;
StartDate:string;
State2:string;
Sustenance_x0020_Total:number;
SustenanceRatio:number;
SustenanceTotal:string;
TestLink:string;
Total:number;
UsuariosCC:string;

}



export interface SPListFA01Records {
  'field%5F0'	:string;//NameFA01
  'field%5F1' :string;//Description
  'field%5F2'	:string;//Area
	'field%5F3':string;//BudgetItem
	'field%5F4':string;//ProjectNumber	
   'Field_5'	:number;//AmountUSD
   'field%5F6':number;//AmountLocal
   'field%5F7'	:Date;//Date
   'field%5F8':string;//Status
   'field%5F9':string;//Region
   'field%5F11':string;//Location
   'field%5F12'	:string;//Project Type
   'field%5F13':string;//CASA Type
   'field%5F15'	:string;//PjctSponsorDisp
   'field%5F16':string;// DesiredDa
   'field%5F31':string;//VoidControl
   'field%5F18':string;//LeaseRequired
   'field%5F20':string;//Budget Item Name
   'field%5F21':number;//BudgetedAmmount
   'field%5F22'	:number;//Budget Next Yr
   'field%5F23':string;//Budget In 2 Yr
   'field%5F24'	:Date; //Start Date
   'field%5F25'	:Date;//EndDate
   'field%5F26'	:string;//Owner
  //Approval Status
  'field%5F17':string;//Void Control
  'field%5F32':string;//Budgeted
  'field%5F33':string;//Company
  'field%5F34':string;//Item
  'field%5F35':string;//ProjectObjectives
  'field%5F36':string;//CASA
  'field%5F37':string;//CompanyNumber
  //Compliance Asset Id	
  'field%5F39':string;//Country
  'field%5F42'	:string;//FA01Description
  //Folder Child Count	:string;
  'field%5F44':string;//Notices Sent
  'field%5F45':string;//NoticesSent
  'field%5F46':string;//Pending Approver
  'field%5F47':string;//Purpose
  'field%5F48':string;//QuotesJustification
  'Title'	:string;//Title
  'field%5F52':string;// Actual Responsible
  'field%5F53':string;//Actual Responsible Disp
  'field%5F56':number;//Amount USD
  'field%5F59':string;// Approval Date
  'field%5F60':string;//Approval1
  'field%5F61':string;//Approval10
  'field%5F62':string;// Approval2
  'field%5F63':string;//Approval3
  'field%5F64':string;//Approval4
  'field%5F65':string;//Approval
  'field%5F66'	:string;//Approval
  'field%5F67':string;//Approval7
  'field%5F68':string;//Approval8
  'field%5F69':string;// Approval9
  'field%5F70':string;// Approver10Comments
  'field%5F72':string;//Approver1Comments
  'field%5F74':string;//Approver2Comments
  'field%5F76':string;//Approver3Comments
  'field%5F78':string;// Approver4Comments
  'field%5F80':string; //Approver5Comments
  'field%5F82':string;//Approver6Comments
  'field%5F84':string;//Approver7Comments
  'field%5F86':string;//Approver8Comments
  'field%5F88':string;//Approver9Comments
  'field%5F90':string;//Aprvr10AprvlDate
  'field%5F91':string;//Aprvr1AprvlDate
  'field%5F92':string // Aprvr2AprvlDate
  'field%5F93':string;//Aprvr3AprvlDate
  'field%5F94':string;//Aprvr4AprvlDate
  'field%5F95':string;//Aprvr5AprvlDate
  'field%5F96':string;//Aprvr6AprvlDate
  'field%5F97':string;//Aprvr7AprvlDate
  'field%5F98':string;//Aprvr8AprvlDate
  'field%5F99':string;//Aprvr9AprvlDate
  'field%5F101':string;//BudgetedJustification
  'field%5F102':string;//BudgetedPast
  'field%5F103':string;//BudgetedPastJustification
  'field%5F104'	:string;//Casa Project
  'field%5F105':string;//CASA Ratio
  'field%5F106':string;//CASA Total
  'field%5F107':string;//CASERatio
  'field%5F108':string;//CASETotal
  //Checked Out To	
  'field%5F115':string;//DFCIncluded
  'field%5F116':string;//DFCJustificacion 
  'field%5F117':string;// DocIncluded
  'field%5F118':string;//DocJustification
  'field%5F119':string;//Expansion Ratio
  'field%5F121':number;//ExpantionRatio
  'field%5F122':number;//ExpantionTotal
  'field%5F123':number;//ExpensesLocal
  'field%5F124':number;//ExpensesUSD
  'field%5F138':string;//GermanyMapping
  'field%5F139':number;//GrandTotal
  //ID	
  'field%5F141':string;//ID2
  'field%5F142':string;//Internal Payback Minimun
  'field%5F143':string;//Is CASA
  'field%5F144':string;//Is cheked out to local
  //Item Child Count	:string;
  //Item Child Count2	:string;
  'field%5F147':string;//Lease User
  'field%5F148':string;//LeaseApproval
  'field%5F149'	:string;//LeaseDisp
  'field%5F150':string;//LeasePending
  'field%5F151'	:string;//LSAprvlDate
  'field%5F152'	:string;//LvBIncluded
  'field%5F153'	:string;//LvBJustification
  'field%5F158'	:string;// Name2
  'field%5F159':string;//Non CASA Total
  'field%5F170'	:string;//ObsLeaseAuth
  'field%5F171':string;// Old Form
  'field%5F172':string;//OUBudget
  'field%5F183'	:number;//Q 1 Budget
  'field%5F184'	:number;//Q 2 Budget
  'field%5F185'	:number;// Q 3 Budget
  'field%5F186'	:number;//Q 4 Budget
  'field%5F187':string;// Quotes
  'field%5F188':string;//RecurringLocal
  'field%5F189f':string;//RecurringUSD
  'field%5F190'	:string;//RejectionJustificatio
  'field%5F191'	:string;// Restart Approval
  'field%5F192':string;//RiskAssessmen
  'field%5F193'	:string;//RiskAssessmentRequired
  'field%5F194'	:string;//StartDate
  'field%5F195':string;//State2
  'field%5F196':number;//Sustenance Total
  'field%5F197'	:number;//SustenanceRatio
  'field%5F198':string;//SustenanceTotal
  'field%5F199':string;//TestLink
  'field%5F200':number;// Total
  'field%5F201'	:string;//UsuariosCC

}
export interface SPLista {
  value: SPListItemSource[];
 }

 export interface SPListaFA01 {
  value: SPListFA01Request[];
 }

  export interface SPListItemSource {

    Title: string;
    Id: string;
    CategoriaCasa: string;
    TituloDeTIC: string;
    Proceso:string;
    Causa:string;
    Subcausa:string;
    SupervisorInvolucrado:string;
    TipoIncidente:string;
    DescripcionDeIncidente:string;
    ClasificacionTOC:string;
    ImpactoDeIncidentes:string;
    OtroImpactoIncidente:string;
    TipoDeTOC:string;
    Sitio:string;
    Compa_x00f1_ia:string;
    Regi_x00f3_n:string;
    ClasificacionIncidente:string;
    EstadoDeTIC:string;
    NaturalezaIncidente:string;
    NoDeTIC:string;
    EstadoProgreso:string;
    QuienReportaTIC:string;
    TICTOC:string;
    AccionesTomadas:string;
    Estado:string;




    Fechadeinicio: Date;
    FechaCompromiso:Date;
    FechaDeCierre:string;

    Recurrente:string;
    EsAuditoria:string;
    Efectividad:string;
    TocDestacada:boolean;
    TicCerrada:boolean;

    
    tmpRecurrent:boolean;
    tmpEsAuditoria:boolean;
    tmpEfectividad:boolean;

    IniciadorDeTICUsername:string;IniciadorDeTICUsernameId:string;
    EmailRespAtencion:string;EmailRespAtencionId:string;
    Nueva_x0020_columna1:string;//UsernameRespVerificacion
    Nueva_x0020_columna1Id:string;
    CC:string;

    ID: number;

    
  }

const allUsers = [
  {"Id":152,"Email":"amtorres@brenntagla.com"},{"Id": 408,"Email":"alopes@brenntagla.com"},{"Id": 187,"Email":"areis@brenntagla.com"},{"Id": 267,"Email":"afraustro@brenntagla.com"},{"Id": 313,"Email":"apimenta@brenntagla.com"},{"Id": 246,"Email":"apollola@brenntagla.com"},{"Id": 137,"Email":"Aavalos@brenntagla.com"},{"Id": 241,"Email":"amrodriguez@brenntagla.com"},{"Id": 209,"Email":"agallini@brenntagla.com"},{"Id": 362,"Email":"alpena@brenntagla.com"},{"Id": 177,"Email":"avalenzuela@brenntagla.com"},{"Id": 165,"Email":"ahermans@brenntagla.com"},{"Id": 183,"Email":"alramirez@brenntagla.com"},{"Id": 337,"Email":"asabino@brenntagla.com"},{"Id": 273,"Email":"atrivino@brenntagla.com"},{"Id": 97,"Email":"avillalva@brenntagla.com"},{"Id": 259,"Email":"aservellon@brenntagla.com"},{"Id": 240,"Email":"aangeles@brenntagla.com"},{"Id": 743,"Email":"aalves@brenntagla.com"},{"Id": 130,"Email":"agreco@brenntagla.com"},{"Id": 232,"Email":"avelasco@brenntagla.com"},{"Id": 25,"Email":"asakiama@brenntagla.com"},{"Id": 247,"Email":"ajuarez@brenntagla.com"},{"Id": 242,"Email":"alrodriguez@brenntagla.com"},{"Id": 136,"Email":"atrevino@brenntagla.com"},{"Id": 208,"Email":"acastro@brenntagla.com"},{"Id": 399,"Email":"amartins@brenntagla.com"},{"Id": 116,"Email":"arossato@brenntagla.com"},{"Id": 55,"Email":"aklauck@brenntagla.com"},{"Id": 260,"Email":"achaparro@brenntagla.com"},{"Id": 145,"Email":"aquitian@brenntagla.com"},{"Id": 121,"Email":"atriana@brenntagla.com"},{"Id": 371,"Email":"achaves@brenntagla.com"},{"Id": 72,"Email":"acisneros@brenntagla.com"},{"Id": 62,"Email":"aortega@brenntagla.com"},{"Id": 23,"Email":"ansanchez@brenntagla.com"},{"Id": 105,"Email":"zf.atorres@brenntagla.com"},{"Id": 117,"Email":"aclaros@brenntagla.com"},{"Id": 48,"Email":"bguzman@brenntagla.com"},{"Id": 46,"Email":"achacin@brenntagla.com"},{"Id": 143,"Email":"cagonzalez@brenntagla.com"},{"Id": 224,"Email":"avanegas@brenntagla.com"},{"Id": 60,"Email":"asesor.arl@brenntagla.com"},{"Id": 98,"Email":"asesor.arlmosq@brenntagla.com"},{"Id": 167,"Email":"acandelaria@brenntagla.com"},{"Id": 377,"Email":"arivera@brenntagla.com"},{"Id": 390,"Email":"bagomez@brenntagla.com"},{"Id": 345,"Email":"bplaza@brenntagla.com"},{"Id": 361,"Email":"beberle@brenntagla.com"},{"Id": 276,"Email":"bverdesoto@brenntagla.com"},{"Id": 304,"Email":"bosornio@brenntagla.com"},{"Id": 317,"Email":"cangonzalez@brenntagla.com"},{"Id": 149,"Email":"ccaiminagua@brenntagla.com"},{"Id": 205,"Email":"cleyton@brenntagla.com"},{"Id": 302,"Email":"cramirez@brenntagla.com"},{"Id": 89,"Email":"ccerda@brenntagla.com"},{"Id": 22,"Email":"cchigne@brenntagla.com"},{"Id": 91,"Email":"cguaqueta@brenntagla.com"},{"Id": 403,"Email":"caguzman@brenntagla.com"},{"Id": 114,"Email":"chsilva@brenntagla.com"},{"Id": 57,"Email":"coliveira@brenntagla.com"},{"Id": 17,"Email":"carlosrodriguez@brenntagla.com"},{"Id": 274,"Email":"csandoval@brenntagla.com"},{"Id": 412,"Email":"coviedo@brenntagla.com"},{"Id": 291,"Email":"candaluz@brenntagla.com"},{"Id": 173,"Email":"cgutierrez@brenntagla.com"},{"Id": 56,"Email":"cquintana@brenntagla.com"},{"Id": 172,"Email":"cbarrera@brenntagla.com"},{"Id": 182,"Email":"cecheverri@brenntagla.com"},{"Id": 18,"Email":"cperilla@brenntagla.com"},{"Id": 253,"Email":"cbischof@brenntagla.com"},{"Id": 179,"Email":"cvaldivia@brenntagla.com"},{"Id": 314,"Email":"chuayna@brenntagla.com"},{"Id": 196,"Email":"calmarza@brenntagla.com"},{"Id": 185,"Email":"cgonzalez@brenntagla.com"},{"Id": 200,"Email":"ccisneros@brenntagla.com"},{"Id": 20,"Email":"carteaga@brenntagla.com"},{"Id": 278,"Email":"ccarvajal@brenntagla.com"},{"Id": 742,"Email":"croliveira@brenntagla.com"},{"Id": 141,"Email":"ccardoso@brenntagla.com"},{"Id": 407,"Email":"cferreira@brenntagla.com"},{"Id": 249,"Email":"calbrecht@brenntagla.com"},{"Id": 100,"Email":"ccanastro@brenntagla.com"},{"Id": 336,"Email":"cfreitas@brenntagla.com"},{"Id": 231,"Email":"cmartinez@brenntagla.com"},{"Id": 368,"Email":"cruiz@brenntagla.com"},{"Id": 393,"Email":"dlopezg@brenntagla.com"},{"Id": 151,"Email":"dfigueiredo@brenntagla.com"},{"Id": 104,"Email":"dcortez@brenntagla.com"},{"Id": 215,"Email":"dmaya@brenntagla.com"},{"Id": 303,"Email":"dquiros@brenntagla.com"},{"Id": 180,"Email":"ddiaz@brenntagla.com"},{"Id": 254,"Email":"dherrera@brenntagla.com"},{"Id": 415,"Email":"dsamarro@brenntagla.com"},{"Id": 320,"Email":"dsibilin@brenntagla.com"},{"Id": 35,"Email":"dperdomo@brenntagla.com"},{"Id": 740,"Email":"dsouza@brenntagla.com"},{"Id": 175,"Email":"djoliveira@brenntagla.com"},{"Id": 344,"Email":"dbenitez@brenntagla.com"},{"Id": 735,"Email":"dalba@brenntagla.com"},{"Id": 255,"Email":"dcruz@brenntagla.com"},{"Id": 405,"Email":"damezcua@brenntagla.com"},{"Id": 307,"Email":"dpierotty@brenntagla.com"},{"Id": 125,"Email":"dispensario.agricola@brenntagla.com"},{"Id": 119,"Email":"ddavila@brenntagla.com"},{"Id": 225,"Email":"emoreno@brenntagla.com"},{"Id": 140,"Email":"etovar@brenntagla.com"},{"Id": 170,"Email":"edhernandez@brenntagla.com"},{"Id": 198,"Email":"esanchez@brenntagla.com"},{"Id": 42,"Email":"eamorim@brenntagla.com"},{"Id": 134,"Email":"edgarcia@brenntagla.com"},{"Id": 355,"Email":"ezapata@brenntagla.com"},{"Id": 294,"Email":"eguevara@brenntagla.com"},{"Id": 311,"Email":"eherrlein@brenntagla.com"},{"Id": 295,"Email":"emgutierrez@brenntagla.com"},{"Id": 222,"Email":"eali@brenntagla.com"},{"Id": 32,"Email":"emromero@brenntagla.com"},{"Id": 111,"Email":"enfermeria.ec@brenntagla.com"},{"Id": 244,"Email":"ehoyos@brenntagla.com"},{"Id": 243,"Email":"ekgarcia@brenntagla.com"},{"Id": 346,"Email":"emagana@brenntagla.com"},{"Id": 101,"Email":"gdiaz@brenntagla.com"},{"Id": 319,"Email":"emendoza@brenntagla.com"},{"Id": 385,"Email":"ebalboa@brenntagla.com"},{"Id": 160,"Email":"enitzke@brenntagla.com"},{"Id": 193,"Email":"eserrano@brenntagla.com"},{"Id": 402,"Email":"ezilio@brenntagla.com"},{"Id": 328,"Email":"ezsouza@brenntagla.com"},{"Id": 375,"Email":"fchinchilla@brenntagla.com"},{"Id": 213,"Email":"fperara@brenntagla.com"},{"Id": 289,"Email":"fcadena@brenntagla.com"},{"Id": 327,"Email":"fdeleon@brenntagla.com"},{"Id": 41,"Email":"fpvillanueva@brenntagla.com"},{"Id": 212,"Email":"fheinert@brenntagla.com"},{"Id": 38,"Email":"flaverde@brenntagla.com"},{"Id": 108,"Email":"ftobias@brenntagla.com"},{"Id": 210,"Email":"fcargua@brenntagla.com"},{"Id": 166,"Email":"fchinchayan@brenntagla.com"},{"Id": 357,"Email":"fheevel@brenntagla.com"},{"Id": 148,"Email":"fenascimento@brenntagla.com"},{"Id": 174,"Email":"fortega@brenntagla.com"},{"Id": 340,"Email":"feortiz@brenntagla.com"},{"Id": 364,"Email":"fvisconti@brenntagla.com"},{"Id": 223,"Email":"fpisetta@brenntagla.com"},{"Id": 270,"Email":"fbahamondes@brenntagla.com"},{"Id": 391,"Email":"fastudillo@brenntagla.com"},{"Id": 234,"Email":"fgaguilera@brenntagla.com"},{"Id": 372,"Email":"garguedas@brenntagla.com"},{"Id": 354,"Email":"ggonzales@brenntagla.com"},{"Id": 395,"Email":"ghidalgo@brenntagla.com"},{"Id": 383,"Email":"gmachado@brenntagla.com"},{"Id": 164,"Email":"gpozos@brenntagla.com"},{"Id": 380,"Email":"gaguilar@brenntagla.com"},{"Id": 382,"Email":"gservin@brenntagla.com"},{"Id": 68,"Email":"gavasconcelos@brenntagla.com"},{"Id": 217,"Email":"gpluas@brenntagla.com"},{"Id": 251,"Email":"gvera@brenntagla.com"},{"Id": 37,"Email":"gacosta@brenntagla.com"},{"Id": 82,"Email":"ghernandez@brenntagla.com"},{"Id": 411,"Email":"gaviles@brenntagla.com"},{"Id": 159,"Email":"gbenetti@brenntagla.com"},{"Id": 401,"Email":"gsani@brenntagla.com"},{"Id": 339,"Email":"giaguilar@brenntagla.com"},{"Id": 63,"Email":"gchavez@brenntagla.com"},{"Id": 163,"Email":"gcalderon@brenntagla.com"},{"Id": 138,"Email":"gcabilla@brenntagla.com"},{"Id": 84,"Email":"gperez@brenntagla.com"},{"Id": 118,"Email":"hsvargas@brenntagla.com"},{"Id": 315,"Email":"htrabucco@brenntagla.com"},{"Id": 44,"Email":"hnavarrete@brenntagla.com"},{"Id": 176,"Email":"hiflores@brenntagla.com"},{"Id": 146,"Email":"hbillordo@brenntagla.com"},{"Id": 370,"Email":"isola@brenntagla.com"},{"Id": 45,"Email":"imesa@brenntagla.com"},{"Id": 400,"Email":"iarodriguez@brenntagla.com"},{"Id": 120,"Email":"igavarrete@brenntagla.com"},{"Id": 88,"Email":"izapatal@brenntagla.com"},{"Id": 203,"Email":"ihernandez@brenntagla.com"},{"Id": 734,"Email":"jagomez@brenntagla.com"},{"Id": 110,"Email":"jmeneses@brenntagla.com"},{"Id": 236,"Email":"jdmitrieva@brenntagla.com"},{"Id": 64,"Email":"zf.jjordan@brenntagla.com"},{"Id": 266,"Email":"jssantos@brenntagla.com"},{"Id": 49,"Email":"jaruiz@brenntagla.com"},{"Id": 418,"Email":"jchavarria@brenntagla.com"},{"Id": 358,"Email":"jijimenez@brenntagla.com"},{"Id": 150,"Email":"jmendoza@brenntagla.com"},{"Id": 190,"Email":"jspagna@brenntagla.com"},{"Id": 153,"Email":"jvela@brenntagla.com"},{"Id": 316,"Email":"jcarrasco@brenntagla.com"},{"Id": 287,"Email":"jjuarez@brenntagla.com"},{"Id": 147,"Email":"jecabrera@brenntagla.com"},{"Id": 204,"Email":"jvelasquez@brenntagla.com"},{"Id": 39,"Email":"jbarrios@brenntagla.com"},{"Id": 335,"Email":"jribeiro@brenntagla.com"},{"Id": 334,"Email":"jpaulo@brenntagla.com"},{"Id": 51,"Email":"jchino@brenntagla.com"},{"Id": 221,"Email":"jmaqueda@brenntagla.com"},{"Id": 409,"Email":"jburgos@brenntagla.com"},{"Id": 40,"Email":"johernandez@brenntagla.com"},{"Id": 387,"Email":"jlemus@brenntagla.com"},{"Id": 363,"Email":"jmadariaga@brenntagla.com"},{"Id": 374,"Email":"jchinchilla@brenntagla.com"},{"Id": 272,"Email":"jjanin@brenntagla.com"},{"Id": 16,"Email":"jabenitez@brenntagla.com"},{"Id": 324,"Email":"jorgomez@brenntagla.com"},{"Id": 195,"Email":"jmarquez@brenntagla.com"},{"Id": 206,"Email":"jomontero@brenntagla.com"},{"Id": 343,"Email":"janaya@brenntagla.com"},{"Id": 414,"Email":"jgalaviz@brenntagla.com"},{"Id": 135,"Email":"jhernandez@brenntagla.com"},{"Id": 369,"Email":"jjmora@brenntagla.com"},{"Id": 81,"Email":"jascencion@brenntagla.com"},{"Id": 139,"Email":"jgonzalez@brenntagla.com"},{"Id": 297,"Email":"jmdelossantos@brenntagla.com"},{"Id": 306,"Email":"oacosta@brenntagla.com"},{"Id": 342,"Email":"jvalenzuela@brenntagla.com"},{"Id": 341,"Email":"jschork@brenntagla.com"},{"Id": 360,"Email":"jalves@brenntagla.com"},{"Id": 226,"Email":"jolguin@brenntagla.com"},{"Id": 53,"Email":"jcmartinez@brenntagla.com"},{"Id": 197,"Email":"jmedina@brenntagla.com"},{"Id": 376,"Email":"jsotelo@brenntagla.com"},{"Id": 207,"Email":"jdoldan@brenntagla.com"},{"Id": 78,"Email":"juannarvaez@brenntagla.com"},{"Id": 310,"Email":"jlaguna@brenntagla.com"},{"Id": 356,"Email":"jalva@brenntagla.com"},{"Id": 220,"Email":"jviejo@brenntagla.com"},{"Id": 36,"Email":"jbaez@brenntagla.com"},{"Id": 398,"Email":"jpita@brenntagla.com"},{"Id": 34,"Email":"jdiaz@brenntagla.com"},{"Id": 257,"Email":"kaguilera@brenntagla.com"},{"Id": 142,"Email":"kjimenez@brenntagla.com"},{"Id": 299,"Email":"kleiva@brenntagla.com"},{"Id": 161,"Email":"kgonzalez@brenntagla.com"},{"Id": 31,"Email":"kmadrigal@brenntagla.com"},{"Id": 288,"Email":"krodriguez@brenntagla.com"},{"Id": 261,"Email":"kpesantes@brenntagla.com"},{"Id": 282,"Email":"kwong@brenntagla.com"},{"Id": 373,"Email":"kmarti@brenntagla.com"},{"Id": 301,"Email":"laboratoriochile@brenntagla.com"},{"Id": 739,"Email":"lbedoya@brenntagla.com"},{"Id": 365,"Email":"lney@brenntagla.com"},{"Id": 192,"Email":"lbartolacci@brenntagla.com"},{"Id": 43,"Email":"lmoreira@brenntagla.com"},{"Id": 113,"Email":"lmaradiaga@brenntagla.com"},{"Id": 275,"Email":"lvargas@brenntagla.com"},{"Id": 230,"Email":"lcastellanos@brenntagla.com"},{"Id": 87,"Email":"lgil@conquimica.com"},{"Id": 281,"Email":"lfujimoto@brenntagla.com"},{"Id": 419,"Email":"lilopez@brenntagla.com"},{"Id": 312,"Email":"ldiez@brenntagla.com"},{"Id": 124,"Email":"lguzman@brenntagla.com"},{"Id": 406,"Email":"movalle@brenntagla.com"},{"Id": 292,"Email":"lfernandez@brenntagla.com"},{"Id": 298,"Email":"lcortes@brenntagla.com"},{"Id": 290,"Email":"licastro@brenntagla.com"},{"Id": 410,"Email":"lortellado@brenntagla.com"},{"Id": 65,"Email":"lsaldano@brenntagla.com"},{"Id": 50,"Email":"lbandala@brenntagla.com"},{"Id": 131,"Email":"lamador@brenntagla.com"},{"Id": 293,"Email":"ltravieso@brenntagla.com"},{"Id": 28,"Email":"lcadena@brenntagla.com"},{"Id": 86,"Email":"lujimenez@brenntagla.com"},{"Id": 96,"Email":"lmarino@brenntagla.com"},{"Id": 169,"Email":"mduarte@brenntagla.com"},{"Id": 77,"Email":"Malvika.Kommineni@lumen.com"},{"Id": 744,"Email":"lumen.manoj@brenntagla.com"},{"Id": 388,"Email":"mansanchez@brenntagla.com"},{"Id": 239,"Email":"mvasquez@brenntagla.com"},{"Id": 59,"Email":"mcifuentes@brenntagla.com"},{"Id": 331,"Email":"msegura@brenntagla.com"},{"Id": 186,"Email":"mchavez@brenntagla.com"},{"Id": 420,"Email":"mcavalim@brenntagla.com"},{"Id": 24,"Email":"mlasak@brenntagla.com"},{"Id": 397,"Email":"mvazquez@brenntagla.com"},{"Id": 326,"Email":"mcruzado@brenntagla.com"},{"Id": 256,"Email":"mamartinez@brenntagla.com"},{"Id": 235,"Email":"mpsilva@brenntagla.com"},{"Id": 227,"Email":"msoledispa@brenntagla.com"},{"Id": 168,"Email":"mfunez@brenntagla.com"},{"Id": 162,"Email":"marenalde@brenntagla.com"},{"Id": 144,"Email":"mperalta@brenntagla.com"},{"Id": 349,"Email":"mgaytan@brenntagla.com"},{"Id": 329,"Email":"mghernandez@brenntagla.com"},{"Id": 305,"Email":"mlmorales@brenntagla.com"},{"Id": 94,"Email":"mgarcia@brenntagla.com"},{"Id": 238,"Email":"msanchez@brenntagla.com"},{"Id": 71,"Email":"mviana@brenntagla.com"},{"Id": 286,"Email":"malvarado@brenntagla.com"},{"Id": 338,"Email":"mrsantacruz@brenntagla.com"},{"Id": 33,"Email":"mortega@brenntagla.com"},{"Id": 378,"Email":"mferrareis@brenntagla.com"},{"Id": 95,"Email":"mpinedo@brenntagla.com"},{"Id": 202,"Email":"mdelaroza@brenntagla.com"},{"Id": 93,"Email":"malopez@brenntagla.com"},{"Id": 417,"Email":"marrieta@brenntagla.com"},{"Id": 389,"Email":"moliveira@brenntagla.com"},{"Id": 58,"Email":"mcaceres@brenntagla.com"},{"Id": 384,"Email":"medicocdq@brenntagla.com"},{"Id": 318,"Email":"medicosmo@brenntagla.com"},{"Id": 109,"Email":"pcasa@brenntagla.com"},{"Id": 366,"Email":"tp.mjaramillo@brenntagla.com"},{"Id": 21,"Email":"mquiroga@brenntagla.com"},{"Id": 52,"Email":"mcenteno@brenntagla.com"},{"Id": 233,"Email":"mrestrepo@brenntagla.com"},{"Id": 122,"Email":"mmurillo@brenntagla.com"},{"Id": 379,"Email":"mparij@brenntagla.com"},{"Id": 347,"Email":"mcavalari@brenntagla.com"},{"Id": 47,"Email":"mlozano@brenntagla.com"},{"Id": 392,"Email":"mbourget@brenntagla.com"},{"Id": 211,"Email":"nmolina@brenntagla.com"},{"Id": 330,"Email":"nflores@brenntagla.com"},{"Id": 218,"Email":"nalcantara@brenntagla.com"},{"Id": 250,"Email":"narrighi@brenntagla.com"},{"Id": 283,"Email":"nbarranco@brenntagla.com"},{"Id": 348,"Email":"ngonzalez@brenntagla.com"},{"Id": 123,"Email":"oacuna@brenntagla.com"},{"Id": 284,"Email":"oduarte@brenntagla.com"},{"Id": 157,"Email":"ocajas@brenntagla.com"},{"Id": 296,"Email":"operdomo@brenntagla.com"},{"Id": 181,"Email":"operacionesmx@brenntagla.com"},{"Id": 156,"Email":"operacionesperu@brenntagla.com"},{"Id": 285,"Email":"oacevedo@brenntagla.com"},{"Id": 733,"Email":"oneto@brenntagla.com"},{"Id": 154,"Email":"ovelasquez@brenntagla.com"},{"Id": 308,"Email":"oguarin@brenntagla.com"},{"Id": 323,"Email":"oescobar@brenntagla.com"},{"Id": 90,"Email":"olagos@brenntagla.com"},{"Id": 264,"Email":"oparra@brenntagla.com"},{"Id": 736,"Email":"pcerna@brenntagla.com"},{"Id": 258,"Email":"projas@brenntagla.com"},{"Id": 158,"Email":"pconsiglieri@brenntagla.com"},{"Id": 67,"Email":"palopez@brenntagla.com"},{"Id": 26,"Email":"psandoval@brenntagla.com"},{"Id": 178,"Email":"ptenorio@brenntagla.com"},{"Id": 741,"Email":"pmonteiro@brenntagla.com"},{"Id": 300,"Email":"pperez@brenntagla.com"},{"Id": 99,"Email":"pureta@brenntagla.com"},{"Id": 201,"Email":"pmagnano@brenntagla.com"},{"Id": 265,"Email":"pmassella@brenntagla.com"},{"Id": 155,"Email":"pgarcia-arroba@brenntagla.com"},{"Id": 69,"Email":"rromeiro@brenntagla.com"},{"Id": 396,"Email":"rrufino@brenntagla.com"},{"Id": 280,"Email":"rcardoso@brenntagla.com"},{"Id": 352,"Email":"rribeiro@brenntagla.com"},{"Id": 171,"Email":"rpanimboza@brenntagla.com"},{"Id": 76,"Email":"ravikumar.s@lumen.com"},{"Id": 325,"Email":"rpin@brenntagla.com"},{"Id": 404,"Email":"rbraum@brenntagla.com"},{"Id": 228,"Email":"rpalencia@brenntagla.com"},{"Id": 245,"Email":"rortiz@brenntagla.com"},{"Id": 359,"Email":"riperez@brenntagla.com"},{"Id": 332,"Email":"rpujols@brenntagla.com"},{"Id": 29,"Email":"rlizama@brenntagla.com"},{"Id": 112,"Email":"rmatarezzi@brenntagla.com"},{"Id": 413,"Email":"rmccolaugh@brenntagla.com"},{"Id": 279,"Email":"rsorto@brenntagla.com"},{"Id": 19,"Email":"rmunoz@brenntagla.com"},{"Id": 54,"Email":"rsousa@brenntagla.com"},{"Id": 106,"Email":"rchirinos@brenntagla.com"},{"Id": 27,"Email":"rmora@brenntagla.com"},{"Id": 103,"Email":"rlopez@brenntagla.com"},{"Id": 194,"Email":"rlbalcazar@brenntagla.com"},{"Id": 269,"Email":"rzambrano@brenntagla.com"},{"Id": 129,"Email":"rabarca@brenntagla.com"},{"Id": 322,"Email":"srivera@brenntagla.com"},{"Id": 214,"Email":"saguilera@brenntagla.com"},{"Id": 92,"Email":"smendez@brenntagla.com"},{"Id": 268,"Email":"svicente@brenntagla.com"},{"Id": 277,"Email":"scabezas@brenntagla.com"},{"Id": 189,"Email":"svanella@brenntagla.com"},{"Id": 737,"Email":"snascimento@brenntagla.com"},{"Id": 333,"Email":"sdiaz@brenntagla.com"},{"Id": 184,"Email":"sdominguez@brenntagla.com"},{"Id": 83,"Email":"sromero@brenntagla.com"},{"Id": 66,"Email":"sortiz@brenntagla.com"},{"Id": 7,"Email":"spadmin@brenntagla.com"},{"Id": 11,"Email":"sharepointsupport@brenntagla.com"},{"Id": 10,"Email":"sharepointsupport2@brenntagla.com"},{"Id": 271,"Email":"szonta@brenntagla.com"},{"Id": 219,"Email":"spalacios@brenntagla.com"},{"Id": 350,"Email":"smarino@brenntagla.com"},{"Id": 237,"Email":"svasquez@brenntagla.com"},{"Id": 102,"Email":"stomasi@brenntagla.com"},{"Id": 422,"Email":"trosa@brenntagla.com"},{"Id": 30,"Email":"trojas@brenntagla.com"},{"Id": 85,"Email":"uherrera@brenntagla.com"},{"Id": 188,"Email":"umonroy@brenntagla.com"},{"Id": 416,"Email":"vramirez@brenntagla.com"},{"Id": 126,"Email":"vgrajales@brenntagla.com"},{"Id": 107,"Email":"vfontoura@brenntagla.com"},{"Id": 381,"Email":"verodriguez@brenntagla.com"},{"Id": 367,"Email":"vsosa@brenntagla.com"},{"Id": 263,"Email":"vrogers@brenntagla.com"},{"Id": 309,"Email":"vigilanciacampana@brenntagla.com"},{"Id": 115,"Email":"vigilanciaperu@brenntagla.com"},{"Id": 421,"Email":"violeta.vera@brenntagla.com"},{"Id": 351,"Email":"vdilonardo@brenntagla.com"},{"Id": 127,"Email":"vpereira@brenntagla.com"},{"Id": 252,"Email":"vrodriguez@brenntagla.com"},{"Id": 35,"Email":"wcumar@brenntagla.com"},{"Id": 199,"Email":"wvarela@brenntagla.com"},{"Id": 191,"Email":"wmacario@brenntagla.com"},{"Id": 216,"Email":"wespitia@brenntagla.com"},{"Id": 128,"Email":"wraymondi@brenntagla.com"},{"Id": 262,"Email":"wlinares@brenntagla.com"},{"Id": 70,"Email":"xjurado@brenntagla.com"},{"Id": 738,"Email":"yrincon@brenntagla.com"},{"Id": 321,"Email":"ycordero@brenntagla.com"},{"Id": 229,"Email":"yybenitez@brenntagla.com"},{"Id": 386,"Email":"zgonzalezc@brenntagla.com"},{"Id": 61,"Email":"zsoto@brenntagla.com}"}
]

//allUsers is a collection of all users in the site
//Search on array allUsers , returns Id

function getUserId(allUsers: any[], EmailToFind: string) {

  if (EmailToFind === null || EmailToFind === undefined) {
    return null;
  }


  let foundId = allUsers.filter(x => x.Email === EmailToFind);

  if (foundId.length > 0) {
    return foundId[0].Id;
  } else {  
    return null;
  }


  if (foundId) {
    return foundId;
  } else {
    return -1;
  }

}









declare let _numberCountSaved: 0;
   declare let _numberCountErrors: 0;
  //  declare let _scriptVersion:1232;

  export class ItemsService{
    private _sp: SPFI;
  
    constructor(private context: WebPartContext) {
      this._sp = spfi().using(SPFx(this.context));
    }
  
    //This is going to bring you many details of an element.
    public async getDetailedListElement(list:string, id:number){
      let singleItemQuery = `<View><Query>
        <Where>
           <Eq>
              <FieldRef Name='ID' />
              <Value Type='Counter'>${id}</Value>
           </Eq>
        </Where>
     </Query></View>`;
      let singleElement = await (
        await this._sp.web.lists
          .getByTitle(list)
          .renderListDataAsStream({ ViewXml: singleItemQuery })
      ).Row[0];
      return singleElement;
    }
    
    //This is going to bring you just the data that the pnp method is calling.
    public async getListElement(list:string, id:number){
      let singleElement = await this._sp.web.lists
        .getByTitle(list)
        .items.getById(id)();
  
       return singleElement;
    }


  
  }


  
  // start the app
 // new AdminTS();


export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  
  
 

  private _sp = spfi().using(SPFx(this.context));
  

  public render(): void {



    this.domElement.innerHTML = `
    <section >
      <div class="">
        <img alt="" src="require('./assets/welcome-dark.png')" class="" />
        <h2> (Version 20230202) Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
        <div>Web part test: <strong>${escape(this.properties.test)}</strong></div>
        <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
      </div>
      <input type="button" value="Click to Run Typescript" id="coolbutton"></input>
      <div>
        <label for="FromID">From ID:</label>
        <input type="text" id="FromID" name="FromID" value="46800">
        </div>
        <div>
        <label for="ToID">To ID:</label>
        <input type="text" id="ToID" name="ToID" value="46900">
        </div>
        <div>

        <button onclick="myFunction()">Run</button>
      </div>





                  <div>


                          <label>select date</label>

                          <input type="date" id="date">

                    <button type="button" id="submit"> submit</button>

                    <h1 id="h" style="color: red;"></h1>

                  </div>     
<script>
                  var t = "test" + new Date();
                  this.domElement.querySelector('#submit').addEventListener('click', () => { 
                 
                    myF()
                   })
                
                  function myF(){
                      var value = document.getElementById("date")["value"];
                      
                      document.getElementById("h")["innerText"]=value;
                  }
     </script>


      <div id="spListContainer" />
      <div id="spListContainer2" />

    </section>`;
    
    //this._renderListasDelSitioAsync();

    if (false)
      this._procesarListaTarjetasDeIncidentes();

      this._procesarListaFA01Requests();


    // this._procesarRecordsParaColumnaCCAUX();

    if (false)
      this._procesarRecordsParaColumnaFechaCierre();


    if (false)              
      this._getAlls();


    
    // for (let i = 0; i < 1; i++) {
    //   let myStart = (i * 500);
    //       this._countItems(myStart);
    // }

  }//end render




  //write myFunction
  public myFunction() {
    // Declare variables
    // var inputFrom;
    // inputFrom = document.getElementById("FromID");
    // var inputTo;
    // inputTo = document.getElementById("ToID");
    // var filter;
    // filter = "ID gt " + inputFrom + " and ID lt " + inputTo;
    // alert(filter);

  }

  // private _renderListasDelSitioAsync(): void {
  //   this._getListData()
  //     .then((response) => {
  //       this._renderListasDelSitio(response.value);
  //     })
  //     .catch(() => {});
  // }
  // private _getListData(): Promise<ISPLists> {
  //   return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
  //       return response.json();
  //     })
  //     .catch(() => {});
  // }

  



  private async _getAlls(): Promise<number>{
    
    

    // ...
    this._sp = spfi().using(SPFx(this.context));
      
    // // basic usage
    // const allItems: any[] =  this._sp.web.lists.getByTitle("Records").items.getAll();
    // console.log("1-getAll():" + allItems.toString());

    // // set page size
    // const allItems2: any[] =  this._sp.web.lists.getByTitle("Records").items.getAll(5000);
    // console.log("2-getAll(5000):" +allItems2.length);

    // // use select and top. top will set page size and override the any value passed to getAll
    // const allItems3: any[] =  this._sp.web.lists.getByTitle("Records").items.select("SourceID").top(5000).getAll();
    // console.log(allItems3.length);

    // // we can also use filter as a supported odata operation, but this will likely fail on large lists
    // const allItems4 =   this._sp.web.lists.getByTitle("Records").items.select("SourceID").filter("SourceID le 5000").getAll();
    // console.log("  _getAlls(): - SourceID le 5000 q:" + (await allItems4).length);


    let myString="";
    //40000
    //50000-51000 . vacio
    //51000-55000 . vacio 
    //55000-56000 . vacio
    //56000-57000 . vacio
    //57000-60000 . vacio
    //60000-65000 . vacio
    //65000-70000 . vacio
  for (let iLoop = 41683;  iLoop  < 41683; iLoop++) {

    try{
          //let oneOtem: any = await this._sp.web.lists.getByTitle("Records").items.getById(iLoop)();
          let oneOtem: any = await this._sp.web.lists.getByTitle("Records").items.filter("SourceID gt "+iLoop).getAll();
          
          myString = myString + oneOtem.SourceID + " , ";

          for (let i = 0; i < 1000000; i++) {
            let myStart = (i * 500);
                myStart = myStart + 1;
          }

          
    }
    catch(e){
    

       if ((e.toString().indexOf("429") > -1) || (e.toString().indexOf("406") > -1) ){
        for (let i = 0; i < 1000000; i++) {
          let myStart = (i * 500);
              myStart = myStart + 1;
        }
      }
    }
    }

    console.log("mystring: - :" + myString);
  return 0;

  }

  // private _getListData_Records(desdeId:number,topQItems:number): Promise<SPListRecords> {

  //   const myFilter = "$filter=Id%20gt%20"+ desdeId+ "%20&$top="+topQItems;
    
  //   return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('records')/items?"+myFilter, SPHttpClient.configurations.v1)
  //     .then((response: SPHttpClientResponse) => {
        
  //        return response.json();
  //     });
  //    }
     private _getListData_RecordsConFechaCierre(desdeId:number,topQItems:number): Promise<SPListRecords> {

      // const myFilter = "$filter=Id%20gt%20"+ desdeId+ "%20&$top="+topQItems;
      const myFilter = "$filter=(Id%20gt%20"+ desdeId+ "%20)" + "%20&$top="+topQItems+ "%20&$orderby=Id asc";
      
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('records')/items?"+myFilter, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          
           return response.json();
        });
       }

  private _getListData_tarjetasdeincidentes(desdeId:number): Promise<SPLista> {

    const myFilter = "$filter=Id%20gt%20"+ desdeId+ "%20&$top=500";
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('tarjetasdeincidentes')/items?"+myFilter, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        
         return response.json();
      });

    
  }

  private _getListData_FA01Requests(desdeId:number,myTop:number): Promise<SPListaFA01> {

    const myFilter = "$filter=Id%20gt%20"+ desdeId+ "%20&$top="+myTop+ "%20&$orderby=Id asc"
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('FA01Request')/items?"+myFilter, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {        
         return response.json();
      });

    
  }


  // private _procesarRecordsParaColumnaCCAUX(): void {
  //   //
  //   const topQItems=500;
  //   this._getListData_Records(123372     ,topQItems)
  //   .then((response) => {
  //     this._procesarTodosLosItemsDeRecordsParaCCAUX(response.value);
  //   })
  //   .catch(() => {});

  //   console.log("procesarRecordsParaColumnaCCAUX: Terminado ");
  // }
  public _procesarRecordsParaColumnaFechaCierre(): void {
    //zzz
    const topQItems=500;
    //59983
    let desdeId=0  ;  //sumar 1500 o el ultimo que se vea

    this._getListData_RecordsConFechaCierre(desdeId  ,topQItems)
    .then((response) => {
      this._procesarTodosLosItemsDeRecordsParaFechaCierre(response.value);
    })
    .catch(() => {});

    //based on current datetime wait 2 minutes
     setTimeout(() => {

        console.log("Ya pasaron 1 etapa");
        this._getListData_RecordsConFechaCierre(desdeId +500 ,topQItems)
        .then((response) => {
          this._procesarTodosLosItemsDeRecordsParaFechaCierre(response.value);
        })
        .catch(() => {});

            //based on current datetime wait 2 minutes
            setTimeout(() => {

            console.log("Ya pasaron 2 etapas");
            this._getListData_RecordsConFechaCierre(desdeId +1000 ,topQItems)
            .then((response) => {
              this._procesarTodosLosItemsDeRecordsParaFechaCierre(response.value);
            })
            .catch(() => {});
    
            }, 80000);
      
      }, 80000);

      

 


    console.log("procesarRecordsParaColumnaFechaCierre: Terminado ");
  }
  private _procesarListaTarjetasDeIncidentes(): void {
     this._getListData_tarjetasdeincidentes(99999)
      .then((response) => {
        this._procesarTodosLosItems(response.value);
      })
      .catch(() => {});
  }

  private _procesarListaFA01Requests(): void {
    //yyyyy

   

      this._getListData_FA01Requests(4597,500)
      .then((response) => {
        this._procesarTodosLosItemsFA01(response.value);
      })
      .catch(() => {});

     


 }


  // private _procesarTodosLosItemsDeRecordsParaCCAUX(itemsAProcesar: SPListItemRecord[]): void {
    

  //   itemsAProcesar.forEach((item: SPListItemRecord) => {

  //       // html2 += `<ul class="${styles.list}">
  //       // <li class="${styles.listItem}">
  //       // <span class="ms-font-l">${item.Title} --- ${item.CategoriaCasa} --- ${item.FechaDeInicio}</span>
  //       // </li>
  //       // </ul>`;
  //       // counter++;
  //       // // issue with sharepoint throttling
  //       // // if counter is divsible by 500 then wait 1 minute
        

  //       // if (counter % 500 === 0) {
  //       //              timer(60000).pipe(take(1)).toPromise().then();
  //       // }
  //       {
  //         if (item.CCAUX === null)
  //         {}
  //         else
  //         {
  //         this.updateListItemRecords(item);
  //       }
  //       }
        
        
      
  //   });
  // }

  private _procesarTodosLosItemsDeRecordsParaFechaCierre(itemsAProcesar: SPListItemRecord[]): void {
        itemsAProcesar.forEach((item: SPListItemRecord) => {    
        {
          if (item.fecharealdecierrestring === null)
          {

          }
          else
          {
            this.updateListItemRecordsParaFechaDeCierre(item);
          }
        }
        
        
      
    });
  }//fin de _procesarTodosLosItemsDeRecordsParaFechaCierre

  private _procesarTodosLosItems(itemsAProcesar: SPListItemSource[]): void {
    //eslint-disable-next-line no-console
    //let html2: string = '';
    // let counter:number=0;  

    

    itemsAProcesar.forEach((item: SPListItemSource) => {

        // html2 += `<ul class="${styles.list}">
        // <li class="${styles.listItem}">
        // <span class="ms-font-l">${item.Title} --- ${item.CategoriaCasa} --- ${item.FechaDeInicio}</span>
        // </li>
        // </ul>`;
        // counter++;
        // // issue with sharepoint throttling
        // // if counter is divsible by 500 then wait 1 minute
        

        // if (counter % 500 === 0) {
        //              timer(60000).pipe(take(1)).toPromise().then();
        // }

          this.addListItemToListDestination(item);
        
        
      
    });

  
  //   // const listContainer2: Element = this.domElement.querySelector('#spListContainer2');
  //   // listContainer2.innerHTML = html2 ;

  //   const miMensaje = "Termino addListItems con Grabados:" + _numberCountSaved + " --- Errores: " + _numberCountErrors;
  //     console.log(miMensaje);
  //     alert (miMensaje);
  }
  private _procesarTodosLosItemsFA01(itemsAProcesar: SPListFA01Request[]): void {
    //eslint-disable-next-line no-console
    //let html2: string = '';
    // let counter:number=0;  

    itemsAProcesar.forEach((item: SPListFA01Request) => {

          this.addListItemToListFA01Record(item);
    });

  
  //   // const listContainer2: Element = this.domElement.querySelector('#spListContainer2');
  //   // listContainer2.innerHTML = html2 ;

     const miMensaje = "Termino addListItems con Grabados:" + _numberCountSaved + " --- Errores: " + _numberCountErrors;
       console.log(miMensaje);
  //     alert (miMensaje);
  }
  private  isDate(myDate: Date):boolean {
    return myDate.constructor.toString().indexOf("Date") > -1;
  } 

  
  




  private  addListItemToListDestination( tarjetadeincidenteSource: SPListItemSource): void {
  

    //Data quality

      //tarjetadeincidenteSource.FechaDeCierre es Texto >>>  Fecharealdecierre es date
      const tmpdate = new Date(tarjetadeincidenteSource.FechaDeCierre);

      if (!this.isDate(tmpdate))
      {        
        console.log("MIGRATIONISSUE(6382) - Error en fecha tarjetadeincidenteSource.FechaDeCierre: "  + tarjetadeincidenteSource.FechaDeCierre)
      }

      // tarjetadeincidenteSource.Recurrente es Texto >>> Recurrent es Yes/No
      // string ( Si - No - Blanco ) >>> Yes/No - Default No
      let tmpRecurrent:boolean;

      switch (tarjetadeincidenteSource.Recurrente) {
        case "Si":
          tmpRecurrent=true;
          break;
        case "No":
          tmpRecurrent=false;  
          break;
      
        default:
          tmpRecurrent=false;// porque el default del campo es false
          break;
      }

      // tarjetadeincidenteSource.EsAuditoria es Texto >>> Auditoria es Yes/No
      // string ( Si - No - Blanco ) >>> Yes/No - Default Yes
      let tmpEsAuditoria:boolean;

      switch (tarjetadeincidenteSource.EsAuditoria) {
        case "Si":
          tmpEsAuditoria=true;
          break;
        case "No":
          tmpEsAuditoria=false;  
          break;
      
        default:
          tmpEsAuditoria=true;// porque el default del campo es true
          break;
      }

      // tarjetadeincidenteSource.Efectividad es Texto >>> efectividaddeacciones es Yes/No
      // string ( Si - No - Blanco ) >>> Yes/No - Default No
      let tmpEfectividad:boolean;

      switch (tarjetadeincidenteSource.Efectividad) {
        case "Si":
          tmpEfectividad=true;
          break;
        case "No":
          tmpEfectividad=false;  
          break;
      
        default:
          tmpEfectividad=false;// porque el default del campo es No
          break;
      }

      //EstadoProgreso
      //0-100 dropdown de 10 en 10	es 100 si estado es Cerrado		sino es 0
      let tmpEstadoProgreso:number;
      if (tarjetadeincidenteSource.EstadoProgreso === null || tarjetadeincidenteSource.EstadoProgreso === "")
      {
        if (tarjetadeincidenteSource.EstadoDeTIC === "TIC Cerrada")
        {
          tmpEstadoProgreso=100;
        }
        else
        {
          tmpEstadoProgreso=0;
        }
      }
      else
      {
        tmpEstadoProgreso=parseInt(tarjetadeincidenteSource.EstadoProgreso,10);
      }

      //fecharealdecierre
      let tmpfecharealdecierre:string="";
      if (tarjetadeincidenteSource.FechaDeCierre === null || tarjetadeincidenteSource.FechaDeCierre === "")
      {
        tmpfecharealdecierre="";
      }
      else
      {        
          tmpfecharealdecierre=tarjetadeincidenteSource.FechaDeCierre;
        
      }



      // IniciadorDeTICUsername:string;IniciadorDeTICUsernameId:number;
      // EmailRespAtencion:string;EmailRespAtencionId:number;
      // Nueva_x0020_columna1:string;Nueva_x0020_columna1Id:number;

      // Begin first call and store promise without waiting
      //const resultGetUserId1 = await _getUserId(tarjetadeincidenteSource.IniciadorDeTICUsername,"IniciadorDeTICUsername",this.context)

      // Begin second call and store promise without waiting
      //const resultGetUserId2 = await _getUserId(tarjetadeincidenteSource.EmailRespAtencion,"EmailRespAtencion",this.context)

      // Now we await for both results, whose async processes have already been started
      //const finalResult = [await resultGetUserId1, await resultGetUserId2];

     

      // At this point all calls have been resolved
      // Now when accessing someResult| anotherResult,
      // you will have a value instead of a promise
      // if (finalResult)
      // {
      //   //tarjetadeincidenteSource.IniciadorDeTICUsernameId=resultGetUserId1.result;
      //   //tarjetadeincidenteSource.EmailRespAtencionId=resultGetUserId2;  
      // }

      const siteURL = this.context.pageContext.web.absoluteUrl;
      const listName = 'records';
      const url = `${siteURL}/_api/web/lists/getbyTitle('${listName}')/items`;
      const body: string = JSON.stringify({
  
        //los campos deben respetar el case
        // metadata contains type information with the listname
         '__metadata': { 'type': 'SP.Data.RecordsListItem' },
        //'__metadata': { 'type': 'SP.List' },
        'Title': tarjetadeincidenteSource.Title,
        'CategoriaCasa': tarjetadeincidenteSource.CategoriaCasa,
        'FechaDeInicio': tarjetadeincidenteSource.Fechadeinicio,
        'Titulo':tarjetadeincidenteSource.TituloDeTIC,
        'Proceso':tarjetadeincidenteSource.Proceso,
        'Causa':tarjetadeincidenteSource.Causa,
        'SubCausa':tarjetadeincidenteSource.Subcausa,
        'Supervisoracargo':tarjetadeincidenteSource.SupervisorInvolucrado,
        'Tipodeincidente':tarjetadeincidenteSource.TipoIncidente,
        'Observaciones':tarjetadeincidenteSource.DescripcionDeIncidente,
        'ClasificacionTOC':tarjetadeincidenteSource.ClasificacionTOC,
        'ImpactodeIncidente':tarjetadeincidenteSource.ImpactoDeIncidentes,
        'Otroimpactodeincidente':tarjetadeincidenteSource.OtroImpactoIncidente,
        'TipodeTOC':tarjetadeincidenteSource.TipoDeTOC,
        'Sitio':tarjetadeincidenteSource.Sitio,
        'Compania':tarjetadeincidenteSource.Compa_x00f1_ia,
        'Region':tarjetadeincidenteSource.Regi_x00f3_n,
        'Incidente':tarjetadeincidenteSource.ClasificacionIncidente,
        'EstadodeRevision':tarjetadeincidenteSource.EstadoDeTIC,
        'Naturalezadelincidente':tarjetadeincidenteSource.NaturalezaIncidente,
        'NumerodeTIC':tarjetadeincidenteSource.NoDeTIC,
        'Reportadopor':tarjetadeincidenteSource.QuienReportaTIC,
        'Seleccion':tarjetadeincidenteSource.TICTOC,
        'Soluci_x00f3_nempleada':tarjetadeincidenteSource.AccionesTomadas,
        'Status':tarjetadeincidenteSource.Estado,
  
        'Fechadecompromiso':tarjetadeincidenteSource.FechaCompromiso,
       // 'Fecharealdecierre':tarjetadeincidenteSource.FechaDeCierre,   //esperando dataquality // se migra a fecharealdecierrestring

        'fecharealdecierrestring':tmpfecharealdecierre,   //esperando dataquality

  
        'Recurrente':tmpRecurrent,
        'Auditor_x00ed_a':tmpEsAuditoria,
        'Efectividaddeacciones':tmpEfectividad,
        'tocdestacada':tarjetadeincidenteSource.TocDestacada,
        'toccerrada':tarjetadeincidenteSource.TicCerrada,
  
  
        // AssignedToId: { 'results': [11,22] } 
  
        'IniciadorAUX':tarjetadeincidenteSource.IniciadorDeTICUsername,
        'RespAtencionAUX':tarjetadeincidenteSource.EmailRespAtencion,
        'RespValidacionAUX':tarjetadeincidenteSource.Nueva_x0020_columna1, //RespVerificacion
        'CCAUX':tarjetadeincidenteSource.CC,
        

        //Numbers
        'Subtipodeincidente':tmpEstadoProgreso,
        'SourceID':tarjetadeincidenteSource.ID,
        
        
  
        //xxx
  
  
      })//end body;

      const spHttpClientOptions: ISPHttpClientOptions = {
        body: body,
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        method: 'POST'
      };
  
      console.log("Voy a hacer el post para agregar un item a la lista (1971)");
      this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          console.log(response);        
          //_numberCountSaved++;
        }
        ).catch((error: SPHttpClientResponse) => {
          //_numberCountErrors++; //is not defined
          console.log("MIGRATIONISSUE(3633):Error en addListItems-POST >>> " + error + ">>> (ID)" + tarjetadeincidenteSource.Id);
        });
      

  }//end of addListItemToList

  
  //Recibe un item de la lista FA01Requests y lo agrega a la lista FA01Records
  private  addListItemToListFA01Record( itemSource: SPListFA01Request): void {
    //Data quality
    
    let tmpDate_fecha: Date = fnProcessDate( itemSource.Date);
    let tmpDate_fechaStart: Date = fnProcessDate( itemSource.Start_x0020_Date);
    let tmpFechaEnd: Date = fnProcessDate( itemSource.EndDate);
    let tmpApprovalDate: Date = fnProcessDate( itemSource.Approval_x0020_Date);
    let tmpDateAprvr10AprvlDate: Date = fnProcessDate( itemSource.Aprvr10AprvlDate);
    let tmpDateAprvr1AprvlDate: Date = fnProcessDate( itemSource.Aprvr1AprvlDate);
    let tmpDateAprvr2AprvlDate: Date = fnProcessDate( itemSource.Aprvr2AprvlDate);
    let tmpDateAprvr3AprvlDate: Date = fnProcessDate( itemSource.Aprvr3AprvlDate);
    let tmpDateAprvr4AprvlDate: Date = fnProcessDate( itemSource.Aprvr4AprvlDate);
    let tmpDateAprvr5AprvlDate: Date = fnProcessDate( itemSource.Aprvr5AprvlDate);
    let tmpDateAprvr6AprvlDate: Date = fnProcessDate( itemSource.Aprvr6AprvlDate);
    let tmpDateAprvr7AprvlDate: Date = fnProcessDate( itemSource.Aprvr7AprvlDate);
    let tmpDateAprvr8AprvlDate: Date = fnProcessDate( itemSource.Aprvr8AprvlDate);
    let tmpDateAprvr9AprvlDate: Date = fnProcessDate( itemSource.Aprvr9AprvlDate);
    let tmpDateStartDate: Date = fnProcessDate( itemSource.StartDate);
    let tmpDesiredDate: Date = fnProcessDate( itemSource.DesiredDa);
    

    let tmpUsuariosCC: string = getUserId(allUsers,itemSource.UsuariosCC);
    let tmpApprover1: string = getUserId(allUsers,itemSource.Approver1Comments);
    let tmpApprover2: string = getUserId(allUsers,itemSource.Approver2Comments);
    let tmpApprover3: string = getUserId(allUsers,itemSource.Approver3Comments)
    let tmpApprover4: string = getUserId(allUsers,itemSource.Approver4Comments);
    let tmpApprover5: string = getUserId(allUsers,itemSource.Approver5Comments);
    let tmpApprover6: string = getUserId(allUsers,itemSource.Approver6Comments);
    let tmpApprover7: string = getUserId(allUsers,itemSource.Approver7Comments);
    let tmpApprover8: string = getUserId(allUsers,itemSource.Approver8Comments);
    let tmpApprover9: string = getUserId(allUsers,itemSource.Approver9Comments);
    let tmpApprover10: string = getUserId(allUsers,itemSource.Approver10Comments);



    


      const siteURL = this.context.pageContext.web.absoluteUrl;
      const listName = 'FA01 Records';
      const url = `${siteURL}/_api/web/lists/getbyTitle('${listName}')/items`;
      const body: string = JSON.stringify({
  
        //los campos deben respetar el case
        // metadata contains type information with the listname
         '__metadata': { 'type': 'SP.Data.FA01_x0020_RecordsListItem' },
        //'__metadata': { 'type': 'SP.List' },
        'SourceID': itemSource.Id,
        'Title': itemSource.Title,
        //'field_0'	:itemSource.NameFA01, // Se ignora , es un link a sharepoint on premise
        'field_1' :itemSource.Description,
        'field_2'	:itemSource.Area,
        'field_3':itemSource.BudgetItem,
        'field_4':itemSource.ProjectNumber,
        'Field_5'	:itemSource.AmountUSD,
        'field_6':itemSource.AmountLocal,
        'field_7'	:tmpDate_fecha,
        'field_8':itemSource.Status,
        
        'field_11':itemSource.Location,
        'field_12'	:itemSource.Project_x0020_Type,
        'field_13':itemSource.CASA_x0020_Type,
        // 'field_15'	:itemSource.PjctSponsorDisp,
        'field_16':tmpDesiredDate,
        'field_31':itemSource.VoidControl,
        // 'field_18':itemSource.LeaseRequired,
        'field_20':itemSource.Budget_x0020_Item_x0020_Name,
        'field_21':itemSource.BudgetedAmmount,
        'field_22'	:itemSource.Budget_x0020_Next_x0020_Yr,
        'field_23':itemSource.Budget_x0020_In_x0020_2_x0020_Yr,
        'field_24'	:tmpDate_fechaStart,
        'field_25':tmpFechaEnd,
        'field_26'	:itemSource.Owner,
        //Approval Status
        'field_17':itemSource.VoidControl,
        'field_32':itemSource.Budgeted,
        'field_33':itemSource.Company,
        'field_34':itemSource.Item,
        'field_35':itemSource.ProjectObjectives,
        'field_36':itemSource.CASA,
        'field_37':itemSource.CompanyNumber,
        //Compliance Asset Id

        'field_39':itemSource.Country,
        'field_42':itemSource.FA01Description,
        //Folder Child Count	:string;
        'field_44':itemSource.NoticesSent,
        'field_45':itemSource.NoticesSent,
        'field_46':itemSource.Pending_x0020_Approver,
        'field_47':itemSource.Purpose,
        'field_48':itemSource.QuotesJustification,

        // Pending Approver Disp
        // 'field_52':itemSource.Actual_x0020_Responsible_x0020_Disp,
        // 'field_53':itemSource.Actual_x0020_Responsible,
        'field_56':itemSource.Amount_x0020_USD,
        'field_59':tmpApprovalDate,
        'field_60':itemSource.Approval1,
        'field_61':itemSource.Approval10,
        'field_62':itemSource.Approval2,
        'field_63':itemSource.Approval3,
        'field_64':itemSource.Approval4,
        'field_65':itemSource.Approval,
        // 'field_66':itemSource.Approval6, //pendieng pendiente revisar
        'field_67':itemSource.Approval7,
        'field_68':itemSource.Approval8,
        'field_69':itemSource.Approval9,
        'field_70':itemSource.Approver10Comments,
        'field_72':itemSource.Approver1Comments,
        'field_74':itemSource.Approver2Comments,
        'field_76':itemSource.Approver3Comments,
        'field_78':itemSource.Approver4Comments,
        'field_80':itemSource.Approver5Comments,
        'field_82':itemSource.Approver6Comments,
        'field_84':itemSource.Approver7Comments,
        'field_86':itemSource.Approver8Comments,
        'field_88':itemSource.Approver9Comments,
        'field_90':tmpDateAprvr10AprvlDate,
        'field_91':tmpDateAprvr1AprvlDate,
        'field_92':tmpDateAprvr2AprvlDate,
        'field_93':tmpDateAprvr3AprvlDate,
        'field_94':tmpDateAprvr4AprvlDate,
        'field_95':tmpDateAprvr5AprvlDate,
        'field_96':tmpDateAprvr6AprvlDate,
        'field_97':tmpDateAprvr7AprvlDate,
        'field_98':tmpDateAprvr8AprvlDate,
        'field_99':tmpDateAprvr9AprvlDate,
        'field_101':itemSource.BudgetedJustification,
        'field_102':itemSource.BudgetedPast,
        'field_103':itemSource.BudgetedPastJustification,
        'field_104':itemSource.Casa_x0020_Project,
        'field_105':itemSource.CASA_x0020_Ratio,
        'field_106':itemSource.CASA_x0020_Total,
        'field_107':itemSource.CASERatio,
        'field_108':itemSource.CASETotal,
        //Checked Out To
        'field_115':itemSource.DFCIncluded,
        'field_116':itemSource.DFCJustificacion,
        'field_117':itemSource.DocIncluded,
        'field_118':itemSource.DocJustification,
        'field_119':itemSource.Expansion_x0020_Ratio,
        'field_121':itemSource.ExpantionRatio,
        'field_122':itemSource.ExpantionTotal,
        'field_123':itemSource.ExpensesLocal,
        'field_124':itemSource.ExpensesUSD,
        'field_138':itemSource.GermanyMapping,
        'field_139':itemSource.GrandTotal,
        'field_141':itemSource.ID2,
        'field_142':itemSource.Internal_x0020_Payback_x0020_Minimun,
        'field_143':itemSource.Is_x0020_CASA,
        'field_144':itemSource.Is_x0020_cheked_x0020_out_x0020_to_x0020_local,
        //Item Child Count	:string;
        //Item Child Count2	:string;
        'field_147':itemSource.Lease_x0020_User,
        'field_148':itemSource.LeaseApproval,
        'field_149':itemSource.LeaseDisp,
        'field_150':itemSource.LeasePending,
        'field_151':itemSource.LSAprvlDate,
        'field_152':itemSource.LvBIncluded,
        'field_153':itemSource.LvBJustification,
        'field_158':itemSource.Name2,  
        'field_159':itemSource.Non_x0020_CASA_x0020_Total,
        'field_170':itemSource.ObsLeaseAuth,
        'field_171':itemSource.Old_x0020_Form,
        'field_172':itemSource.OUBudget,
        'field_183':itemSource.Q_x0020_1_x0020_Budget,
        'field_184':itemSource.Q_x0020_2_x0020_Budget,
        'field_185':itemSource.Q_x0020_3_x0020_Budget,
        'field_186':itemSource.Q_x0020_4_x0020_Budget,
        'field_187':itemSource.Quotes,
        'field_188':itemSource.RecurringLocal,
        'field_189':itemSource.RecurringUSD,
        'field_190':itemSource.RejectionJustificatio,
        'field_191':itemSource.Restart_x0020_Approval,
        'field_192':itemSource.RiskAssessmen,
        'field_193':itemSource.RiskAssessmentRequired,
        'field_194':tmpDateStartDate,
        'field_195':itemSource.State2,
        'field_196':itemSource.Sustenance_x0020_Total,
        'field_197':itemSource.SustenanceRatio,
        'field_198':itemSource.SustenanceTotal,
        // 'field_199':itemSource.TestLink,//se ignora
        'field_200':itemSource.Total,
        'field_201':tmpUsuariosCC,
        'Approver1':tmpApprover1,
        'Approver2':tmpApprover2,
        'Approver3':tmpApprover3,
        'Approver4':tmpApprover4,
        'Approver5':tmpApprover5,
        'Approver6':tmpApprover6,
        'Approver7':tmpApprover7,
        'Approver8':tmpApprover8,
        'Approver9':tmpApprover9,
        'Approver10':tmpApprover10




  
      })//end body;

      const spHttpClientOptions: ISPHttpClientOptions = {
        body: body,
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        method: 'POST'
      };
  
      console.log("Voy a hacer el post para agregar un item a la lista (5971)");
      this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          console.log(response);        
          //_numberCountSaved++;
        }
        ).catch((error: SPHttpClientResponse) => {
          //_numberCountErrors++; //is not defined
          console.log("MIGRATIONISSUE(55633):Error en addListItems-POST >>> " + error + ">>> (ID)" + itemSource.Id);
        });
      



      

  

}//end of addListItemToListFA01Record


  
  //20230206
private  async updateListItemRecordsParaFechaDeCierre( recordsItem: SPListItemRecord): Promise<void> {
  
    //Data quality fecha de cierre
    // Se encontro que fecharealdecierre tiene formatos de fecha diferentes
    // DD/MM/YYYY y MM/DD/YYYY


    let tmpFechaCierre: string = recordsItem.fecharealdecierrestring;
    let tmpDate_fecharealdecierre: Date = new Date();

    


      //tmpFechaCierre is a DIV with a class and  a string inside, extract the string
      // <div class="ExternalClass41B180652FC24CB782E5AB88C0AC4F16">12/31/2019</div>
      tmpFechaCierre= tmpFechaCierre.split(">")[1].split("<")[0];
//if tmpFechaCierre contains /
if (tmpFechaCierre.indexOf("/") > -1) {
      //delete any hour:minute:seconds information from  tmpFechaCierre 
      tmpFechaCierre = tmpFechaCierre.split(" ")[0];

      try {
        //tmpFechaCierre is a string with format DD/MM/YYYY, convert it to MM/DD/YYYY
        const [month, day, year] = tmpFechaCierre.split('/');

        // console.log(month); // 👉️ "07"
        // console.log(day); // 👉️ "21"
        // console.log(year); // 👉️ "2024"
        let tmpMonth: number = +month;
        if (tmpMonth>12)
        {
          tmpDate_fecharealdecierre = new Date(+year, +day-1 , +month+1);
          // console.log("MIGRATIONmessage(9274):Revisar >>> " + tmpDate_fecharealdecierre + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
        }
        else
        {
          tmpDate_fecharealdecierre = new Date(+year, +month-1 , +day+1);
          // console.log("MIGRATIONmessage(7824):Revisar >>> " + tmpDate_fecharealdecierre + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
        }

      } catch (error) {
        // tmpFechaCierre is a string with format MM/DD/YYYY, convert it to DD/MM/YYYY
        // convert tmpFechaCierre to DD/MM/YYYY
        console.log("MIGRATIONISSUE(8993):Error en updateListItemRecordsParaFechaDeCierre >>> " + error + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
      }
    }//end of if tmpFechaCierre contains /
    else
    {
      //does not contain /
      //tmpFechaCierre is a DIV with a class and  a string inside, extract the string
      // <div class="ExternalClass41B180652FC24CB782E5AB88C0AC4F16">2019-11-29</div>
      // tmpFechaCierre= tmpFechaCierre.split(">")[1].split("<")[0];

      //delete any hour:minute:seconds information from  tmpFechaCierre 
      // tmpFechaCierre = tmpFechaCierre.split(" ")[0];

      try {
        const [year3, month3, day3] = tmpFechaCierre.split('-');

        // console.log(month); // 👉️ "07"
        // console.log(day); // 👉️ "21"
        // console.log(year); // 👉️ "2024"
        let tmpMonth: number = +month3;
        if (tmpMonth>12)
        {
          tmpDate_fecharealdecierre = new Date(+year3, +day3-1 , +month3+1);
           console.log("MIGRATIONmessage(2174):Revisar >>> " + tmpDate_fecharealdecierre + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
        }
        else
        {
          tmpDate_fecharealdecierre = new Date(+year3, +month3-1 , +day3);
          //  console.log("MIGRATIONmessage(2124):Revisar >>> " + tmpDate_fecharealdecierre + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
        }

      } catch (error) {
        console.log("MIGRATIONISSUE(7793):Error en updateListItemRecordsParaFechaDeCierre >>> " + error + ">>> (ID)" + recordsItem.ID + ">>> Al componer fecha de cierre");
      }

    }

    recordsItem.Fecharealdecierre= tmpDate_fecharealdecierre;

    console.log(  " - Voy a updatear un item a la lista (3726) con el ID " + recordsItem.ID + " y el fechacierre " + recordsItem.Fecharealdecierre);

    //if myUsersIDs is empty, then do not update the CC field
    if(!isNaN(tmpDate_fecharealdecierre.getTime() ))
    {

      this._sp.web.lists.getByTitle("Records")
      //Lookup id entity names are always suffixed with Id
      .items.getById(recordsItem.ID).update({ Fecharealdecierre: recordsItem.Fecharealdecierre })
      .then((response:IItemUpdateResult) => {
        // console.log(response.data);
      }).catch((error) => {
        console.log("MIGRATIONISSUE(6482):Error en updateListItemRecordsParaFechaDeCierre >>> " + error + ">>> (ID)" + recordsItem.ID + ">>> (FechaCierre)" + recordsItem.CC);
      });
    }
    else
    {
      //Fecha con formato incorrecto
      console.log("MIGRATIONISSUE(3459):No se actualizo el campo fecharealdecierre por algun problema de formato en el ID:"+recordsItem.ID + "fecharealdecierrestring:" + recordsItem.fecharealdecierrestring  );
    }
     

      

  }//end of updateListItemRecords
  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // private async getUserIDByEmail(email:string,myContext:WebPartContext):Promise<number>
  // {
  //     // var url = `${this.props.wpContext.pageContext.site.absoluteUrl}/_api/web/siteusers?$filter=Email eq '${email}'`;
  //     var url = `${myContext.pageContext.site.absoluteUrl}/_api/web/siteusers?$filter=Email eq '${email}'`;
  //     // var userData:any = await  QueryAUser.GetQueryData(this.props.wpContext,url);
  //     var userData:any = await  QueryAUser.GetQueryData(myContext,url)
  //     .then((response) => {
  //       console.log("response de GetQueryData:" + response);
  //       return userData.value[0].Id;
  //       return response;
  //     })
  //     .catch((error) => {
  //       console.log("MIGRATIONISSUE(34532):Error en getUserIDByEmail >>> " + error + ">>> (email)" + email);
  //       return 99
  //     });

  //     return 98
  // }
        

}//end of class



function fnProcessDate(tmpFechaParameter: string): Date {

    let tmpDateReturn: Date = new Date();

    if (tmpFechaParameter != null)
      {
      // is a string with formats dd/mm/yyyy or yyyy-mm-dd
      if (tmpFechaParameter.indexOf("/") > -1) {
        // is a string with format DD/MM/YYYY, convert it to MM/DD/YYYY
        const [day, month, year] = tmpFechaParameter.split('/');
        tmpDateReturn = new Date(+year, +month - 1, +day);
      } else {
        // is a string with format YYYY-MM-DD, convert it to MM/DD/YYYY
        const [year, month, day] = tmpFechaParameter.split('-');
        tmpDateReturn = new Date(+year, +month - 1, +day);
      }
    }
    else
    {
      tmpDateReturn = null;
    }
    return tmpDateReturn;
}
// export  class QueryAUser {
//   public static async GetQueryData(context: WebPartContext, url: string) {
//     var deferred = $.Deferred();
//     console.log(`running query '${url}'`);
//     let _nometaOpt: ISPHttpClientOptions = {
//       headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '', 'Content-type': 'application/json;odata=verbose' }
//     };
//     context.spHttpClient.get(url, SPHttpClient.configurations.v1, _nometaOpt).then(
//       (response: SPHttpClientResponse) => {

//         if (response.status == 200) {
//           console.log("got query results");
//           if (response.headers.get("content-type").indexOf("atom") > -1) {
//             console.log("Got xml instead of json on " + url);
//             response.text().then((text:string) => {
//               deferred.resolve(text);
//             });
//           }
//           else {
//             response.json().then((jsondata: JSON) => {
//               deferred.resolve(jsondata);
//             });
//           }
//         }
//         else {
//           console.log("error getting query results!");
//           console.log(response.status);
//         }
//       }, (err: any) => {
//         debugger;
//         console.log("error getting query results!");
//         console.log(err);
//       });
//     return deferred;
//   }
// }











// async function _getUserId(parUsername: string, parNombreCampo:string,parContext:WebPartContext) {
//   // parUsername with @domain.com
//   let myUserID:string="";
//   let myUserIDs:string="";//for multiple users result is [id1,id2,id3]
//   if (parUsername == null || parUsername == "") {
//     return "";
//   }

//   //BRENNTAGLA\zdergal
//   //BRENNTAGLA<br>zdergal
//   if (parUsername.indexOf('@') == -1) {
//     //user without @
//     if (parUsername.indexOf('\\') > -1) {
//       //user with domain
//       //split domain and user
//       const tmpUser = parUsername.split('\\');
//       const tmpDomain = tmpUser[0];
//       const tmpUser2 = tmpUser[1];
//       parUsername = tmpUser2 + "@" + tmpDomain;
//     }
//     if (parUsername.indexOf('<br>') > -1) {
//       //user with domain
//       //split domain and user
//       const tmpUser = parUsername.split('<br>');
//       const tmpDomain = tmpUser[0];
//       const tmpUser2 = tmpUser[1];
//       parUsername = tmpUser2 + "@" + tmpDomain;
//     }
//   }

//   //multiple users
//   if (parUsername.indexOf(';') > -1) 
//   {
//     const myUserNames = parUsername.split(';');
//     for (let i = 0; i < myUserNames.length; i++) {
//       const oneUserId = _getUserId(myUserNames[i], parNombreCampo,parContext);
//       if (await oneUserId != "") {
//         myUserIDs += oneUserId+",";        
//       }
//     }
//     //remove last comma
//     myUserIDs = myUserIDs.substring(0, myUserIDs.length - 1);    
//     return myUserIDs;
//   }

  
//   if (parUsername.indexOf('@') > -1) {

//     const tmpgetUser = "i:0#.f|membership|"+parUsername;
//     const payload: string = JSON.stringify({
//       'logonName': tmpgetUser // i:0#.f|membership|firstname.lastname@contoso.onmicrosoft.com      
//     });
    
//     var postData: ISPHttpClientOptions = {
//       body: payload
//     };
    
//     var endPoint = parContext.pageContext.site.absoluteUrl+"/_api/web/ensureuser";
//     if(isNull(myUserID))
//     {}

//     parContext.spHttpClient.post(endPoint,
//       SPHttpClient.configurations.v1,
//       postData)
//       .then((response: SPHttpClientResponse) => {
//         //get id from response
       
//         response.json().then((results) => {
//           console.log("Resultado de ensureuser: "+results);
//           return results['Id'];

//           myUserID=results['Id'];

         

//         }//end of response.json
        
//         );
//     })//end of this.context.spHttpClient.post;

//   }//end of if (parUsername.indexOf('@') > -1) {
//   else
//   {
//     //parUsername without @domain.com
//     console.log("MIGRATIONISSUE(3738):Error en _getUserId >>> " + " Formato sin @" +  " para campo:" + parNombreCampo +">>> (ID)" + parUsername);
//     return "";
//   }

// }

