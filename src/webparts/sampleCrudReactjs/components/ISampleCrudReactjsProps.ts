import { SPHttpClient,SPHttpClientResponse } from "@microsoft/sp-http";


export interface ISampleCrudReactjsProps {
  description: string;
  ListName:string;
  spHttpClient:SPHttpClient;
  siteURL:string;
}
