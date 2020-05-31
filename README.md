# Azure function event receiver for Sharepoint (written for NodeJS)

This is a fork of original event receiver writen by Sergei Sergeev. I modernized it a bit and added an example for changing properties during synchonous events.

Here are some basics:

```bash
## install azure functions core tools
npm i -g azure-functions-core-tools

## install nrok globally, it's very usefull during the development
npm i -g ngrok 

## run it
ngrok http 7071 --host-header=localhost

## start your dev server 
npm run az

## register your event receiver using pnp
Add-PnPEventReceiver -List "list" -Name "PNPTest1" -Url "https://yourOwnAddress.ngrok.io/api/AzFuncRer" -EventReceiverType ItemAdding -Synchronization synchronous -SequenceNumber 1

```


### Sample on implementing SharePoint Remote Event Receiver with Azure Function written in TypeScript 

Original companion article written by Sergei can be found [here](http://spblog.net/post/2017/09/09/Using-SharePoint-Remote-Event-Receivers-with-Azure-Functions-and-TypeScript).