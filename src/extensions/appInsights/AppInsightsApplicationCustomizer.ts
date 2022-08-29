import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { ApplicationInsights } from '@microsoft/applicationinsights-web'

import * as strings from 'AppInsightsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppInsightsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppInsightsApplicationCustomizerProperties {
  instrumentationKey: string;

}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppInsightsApplicationCustomizer
  extends BaseApplicationCustomizer<IAppInsightsApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const appKey: string = this.properties.instrumentationKey;

    if (appKey) {
      const appInsights = new ApplicationInsights({
        config: {
          instrumentationKey: appKey
        }
      });

      appInsights.loadAppInsights();
      appInsights.trackPageView();
    }

    return Promise.resolve();
  }
}
