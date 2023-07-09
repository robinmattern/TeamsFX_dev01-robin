import { AppConfigurationClient } from "@azure/app-configuration";
import { DefaultAzureCredential, EnvironmentCredential, ManagedIdentityCredential } from "@azure/identity";
import { SecretClient } from "@azure/keyvault-secrets";
import { Config } from "../../model/Config";

export default class Utilities {
  public static async retrieveConfig(): Promise<Config> {
    const client = this.getClient();
    let siteID = "";
    let listID = "";
    let useSearch: boolean = false;
    let searchQuery: string = "";
    try {
      const siteSetting = await client.getConfigurationSetting({ key: "SiteID"});
      siteID = siteSetting.value!;
      const listSetting = await client.getConfigurationSetting({ key: "ListID"});
      listID = listSetting.value!;
      const useSearchSetting = await client.getConfigurationSetting({ key: "UseSearch" });
      useSearch = useSearchSetting.value?.toLowerCase() === "true" ? true : false;
      const searchQuerySetting = await client.getConfigurationSetting({ key: "SearchQuery" });
      searchQuery = searchQuerySetting.value!;
    }
    catch(error) {
      if (siteID === "") {
        //  siteID = process.env.SITE_ID!;
      }
      if (listID === "") {
        //  listID = process.env.LIST_ID!;
      }
    }
    return Promise.resolve({ SiteID: siteID, ListID: listID, UseSearch: useSearch, SearchQuery: searchQuery });
  }

  public static async saveConfig(newConfig: Config) {
    const client = this.getClient();
    if (newConfig.SiteID) {
      await client.setConfigurationSetting({ key: "SiteID", value: newConfig.SiteID });
    }
    if (newConfig.ListID) {
      await client.setConfigurationSetting({ key: "ListID", value: newConfig.ListID });
    }    
    if (newConfig.SearchQuery) {
      await client.setConfigurationSetting({ key: "SearchQuery", value: newConfig.SearchQuery });
    }
    await client.setConfigurationSetting({ key: "UseSearch", value: newConfig.UseSearch.toString() });
  }

  private static getClient(): AppConfigurationClient {
    const connectionString = process.env.AZURE_CONFIG_CONNECTION_STRING!;
    // const credential = new DefaultAzureCredential();
    // console.log(credential);
    // const client = new AppConfigurationClient(connectionString, credential);
    let client: AppConfigurationClient;
    if (process.env.AZURE_CLIENT_SECRET) {
      const credential = new EnvironmentCredential();
      client = new AppConfigurationClient(connectionString, credential);
    }
    else {
      const credential = new ManagedIdentityCredential();
      client = new AppConfigurationClient(connectionString, credential);
    }
    return client;
  }

  public static async getSecret(name: string): Promise<string> {
    const client = this.getSecretClient();
    try {
      const secret = await client.getSecret(name);
      return secret.value!;
    }
    catch (err) {
      console.log(err);
      return "";
    }
  }

  private static getSecretClient(): SecretClient {
    const connectionString = process.env.AZURE_KEYVAULT_CONNECTION_STRING!;
    const credential = new DefaultAzureCredential();
    const client = new SecretClient(connectionString, credential);
    
    return client;
  }
}