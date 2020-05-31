"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.run = void 0;
const xml2js_1 = require("xml2js");
const fs = require("fs");
const sp_commonjs_1 = require("@pnp/sp-commonjs");
const nodejs_commonjs_1 = require("@pnp/nodejs-commonjs");
function run(context, req) {
    context.log("Running Remote Event Receiver from Azure Function!");
    execute(context, req)
        .catch((err) => {
        console.log(err);
        context.done();
    });
}
exports.run = run;
function execute(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        let soap_string = req.headers['x-sp-errormessage'];
        let sub_soap_string = soap_string.substring(3);
        let data = yield xml2Json(sub_soap_string);
        if (data["s:Envelope"]["s:Body"].ProcessOneWayEvent) {
            yield processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
        }
        else if (data["s:Envelope"]["s:Body"].ProcessEvent) {
            yield processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context);
        }
        else {
            throw new Error("Unable to resolve event type");
        }
    });
}
//sync ItemAdding
function processEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        // for demo: cancel sync -ing RER with error:
        let body = fs.readFileSync('AzFuncRer/rsponse.xml').toString();
        context.res = {
            status: 200,
            headers: {
                "Content-Type": "text/xml"
            },
            body: body,
            isRaw: true
        };
        context.done();
    });
}
//async ItemAdded
function processOneWayEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        let contextToken = eventProperties.ContextToken;
        let itemProperties = eventProperties.ItemEventProperties;
        let creds = {
            clientId: getAppSetting('ClientId'),
            clientSecret: getAppSetting('ClientSecret')
        };
        sp_commonjs_1.sp.setup({
            sp: {
                fetchClientFactory: () => {
                    return new nodejs_commonjs_1.SPFetchClient(itemProperties.WebUrl, creds.clientId, creds.clientSecret);
                },
            },
        });
        let itemUpdate = yield sp_commonjs_1.sp.web.lists.getById(itemProperties.ListId).items.getById(itemProperties.ListItemId)
            .update({
            Title: "VSYS-" + itemProperties.ListItemId
        });
        console.log(itemUpdate);
        context.res = {
            status: 200,
            body: ''
        };
        context.done();
    });
}
function xml2Json(input) {
    return __awaiter(this, void 0, void 0, function* () {
        return new Promise((resolve, reject) => {
            let parser = new xml2js_1.Parser({
                explicitArray: false
            });
            parser.parseString(input, (jsError, jsResult) => {
                if (jsError) {
                    reject(jsError);
                }
                else {
                    resolve(jsResult);
                }
            });
        });
    });
}
function getAppSetting(name) {
    return process.env[name];
}
//# sourceMappingURL=index.js.map