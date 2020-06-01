import { Parser } from "xml2js";
import * as fs from 'fs';
import { sp } from "@pnp/sp-commonjs";
import { SPFetchClient } from "@pnp/nodejs-commonjs";

declare var global: any;

export function run(context: any, req: any): void {
    context.log("Running Remote Event Receiver from Azure Function!");

    execute(context, req)
        .catch((err: any) => {
            console.log(err);
            context.done();
        });
}

// Parses soap request comming from sharepoint and checks whether 
// ProcessOneWayEvent (async) or ProcessEvent (sync) should be run.
async function execute(context: any, req: any): Promise<any> {
    let soap_string : string = req.headers['x-sp-errormessage'];
    let sub_soap_string : string = soap_string.substring(3);
    
    let data = await xml2Json(sub_soap_string); 

    if(data["s:Envelope"]["s:Body"].ProcessOneWayEvent){
        await processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
    } else if(data["s:Envelope"]["s:Body"].ProcessEvent){
        await processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context)
    } else {
        throw new Error("Unable to resolve event type");
    }
}

//sync ItemAdding
async function processEvent(eventProperties: any, context: any): Promise<any> {
    
    // works for -ing events, changeTitle.xml contains envelope which changes Item Title to "Changed on the go!"
    // this can be also chnged to stopWithError which will prevent saving of the item with error
    let body = fs.readFileSync('changeTitle.xml').toString();
    context.res = {
        status: 200,
        headers: {
            "Content-Type": "text/xml" 
        },
        body: body,
        isRaw: true
    } as any;

    context.done(); 
}

//async ItemAdded
async function processOneWayEvent(eventProperties: any, context: any): Promise<any> {
    let contextToken = eventProperties.ContextToken;
    let itemProperties = eventProperties.ItemEventProperties;
    
    let creds: any = {
        clientId: getAppSetting('ClientId'),
        clientSecret: getAppSetting('ClientSecret')
    };

    sp.setup({
        sp:{
            fetchClientFactory: ()=>{
                return new SPFetchClient(itemProperties.WebUrl, creds.clientId, creds.clientSecret);
            },
        },
    });

    let itemUpdate = await sp.web.lists.getById(itemProperties.ListId).items.getById(itemProperties.ListItemId)
        .update({
            Title: "ID of this item is -"+itemProperties.ListItemId
        });

    console.log(itemUpdate);
    
    context.res = {
        status: 200,
        body: ''
    } as any;

    context.done();
}

async function xml2Json(input: string): Promise<any> {
    return new Promise((resolve, reject) => {
        let parser = new Parser({
            explicitArray: false
        });

        parser.parseString(input, (jsError: any, jsResult: any) => {
            if (jsError) {
                reject(jsError);
            } else {
                resolve(jsResult);
            }
        });
    });
}

function getAppSetting(name: string): string {
    return process.env[name] as string;
}