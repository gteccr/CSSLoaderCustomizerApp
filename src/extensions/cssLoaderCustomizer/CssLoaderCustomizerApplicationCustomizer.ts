import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CssLoaderCustomizerApplicationStrings';

const LOG_SOURCE: string = '[CSS Loader Customizer Application]';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICssLoaderCustomizerApplicationProperties {
  // This is an example; replace with your own property
  testMessage: string;
  cssFileName:string;
  cssInternalLocation:string;
  
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CssLoaderCustomizerApplication 
  extends BaseApplicationCustomizer<ICssLoaderCustomizerApplicationProperties> {
  @override
  public onInit(): Promise<void> {
    try{
        Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
        const CSSINTERNALLOCATION = this.properties.cssInternalLocation;
        const publicCDNURL = this.context.pageContext.legacyPageContext.publicCdnBaseUrl;
        const hostName = window.location.host;
        const cssFileURL = `${publicCDNURL}/${hostName}/${CSSINTERNALLOCATION}/${this.properties.cssFileName}`;
        console.log (`${LOG_SOURCE}: CSS File URL ${cssFileURL} `);
        Log.info(LOG_SOURCE,`CSS File URL ${cssFileURL}`);
        
        if(cssFileURL){
          let head: HTMLHeadElement = document.getElementsByTagName("head")[0] || document.documentElement;
          let customStyle: HTMLLinkElement = document.createElement("link");
          customStyle.href = cssFileURL;
          customStyle.rel = "stylesheet";
          customStyle.type = "text/css";
          head.insertAdjacentElement("beforeend", customStyle);
        }
        else{
          console.info(`${LOG_SOURCE}: No file was found`);      
        }
        
        return Promise.resolve();
    }
    catch(exception){
      console.info(`${LOG_SOURCE}: ${exception}`);
      throw new exception (exception)
    }    
  }
}
