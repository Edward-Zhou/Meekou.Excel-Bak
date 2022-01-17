export class AppConsts {
  static remoteServiceBaseUrl: string = "https://localhost:44311";
  static appBaseUrl: string = "https://localhost:3000";
  static appBaseHref: string; // returns angular's base-href parameter value if used during the publish

  static localeMappings: any = [];

  static readonly userManagement = {
    defaultAdminUserName: "admin",
  };

  static readonly localization = {
    defaultLocalizationSourceName: "Mango",
  };

  static readonly authorization = {
    encryptedAuthTokenName: "enc_auth_token",
  };
}
