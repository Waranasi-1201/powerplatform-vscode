/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 */

import jwt_decode from 'jwt-decode';
import { npsAuthentication } from "../common/authenticationProvider";
import {SurveyConstants, httpMethod, queryParameters} from '../common/constants';
import fetch,{RequestInit} from 'node-fetch'
import WebExtensionContext from '../WebExtensionContext';
import { telemetryEventNames } from '../telemetry/constants';
import {getCurrentDataBoundary} from '../utilities/dataBoundary';

export class NPSService{
    public static getCesHeader(accessToken: string) {
        return {
            authorization: "Bearer " + accessToken,
            Accept: 'application/json',
           'Content-Type': 'application/json',
        };
    }

    public static getNpsSurveyEndpoint(): string{
        const region = WebExtensionContext.urlParametersMap?.get(queryParameters.REGION)?.toLowerCase();
        const dataBoundary = getCurrentDataBoundary();
        let npsSurveyEndpoint = '';
        switch (region) {
          case 'tie':
          case 'test':
          case 'preprod':
            switch (dataBoundary) {
              case 'eu':
                npsSurveyEndpoint = 'https://europe.tip1.ces.microsoftcloud.com';
                break;
              default:
                npsSurveyEndpoint = 'https://world.tip1.ces.microsoftcloud.com';
            }
            break;
          case 'prod':
          case 'preview':
            switch (dataBoundary) {
              case 'eu':
                npsSurveyEndpoint = 'https://europe.ces.microsoftcloud.com';
                break;
              default:
                npsSurveyEndpoint = 'https://world.ces.microsoftcloud.com';
            }
            break;
          case 'gov':
          case 'high':
          case 'dod':
          case 'mooncake':
            npsSurveyEndpoint = 'https://world.ces.microsoftcloud.com';
            break;
          case 'ex':
          case 'rx':
          default:
            break;
        }
      
        return npsSurveyEndpoint;
    }

    public static async setEligibility()  {    
        try{
               
                const baseApiUrl = this.getNpsSurveyEndpoint();
                const accessToken: string = await npsAuthentication(SurveyConstants.AUTHORIZATION_ENDPOINT);
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const parsedToken = jwt_decode(accessToken) as any;
                WebExtensionContext.setUserId(parsedToken?.oid)
                const apiEndpoint = `${baseApiUrl}/api/v1/${SurveyConstants.TEAM_NAME}/Eligibilities/${SurveyConstants.SURVEY_NAME}?userId=${parsedToken?.oid}&eventName=${SurveyConstants.EVENT_NAME}&tenantId=${parsedToken.tid}`;
                const requestInitPost: RequestInit = {
                    method: httpMethod.POST,
                    body:'{}',
                    headers:NPSService.getCesHeader(accessToken)
                };
                const requestSentAtTime = new Date().getTime();
                const response = await fetch(apiEndpoint, requestInitPost);
                const result = await response?.json();
                if( result?.Eligibility){
                    WebExtensionContext.telemetry.sendAPISuccessTelemetry(telemetryEventNames.NPS_USER_ELIGIBLE, "NPS Api",httpMethod.POST,new Date().getTime() - requestSentAtTime);
                    WebExtensionContext.setNPSEligibility(true);
                    WebExtensionContext.setFormsProEligibilityId(result?.FormsProEligibilityId);
                }
        }catch(error){
            WebExtensionContext.telemetry.sendErrorTelemetry(telemetryEventNames.NPS_API_FAILED, (error as Error)?.message);
        }
    }
}