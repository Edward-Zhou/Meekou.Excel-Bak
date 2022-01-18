import { AppConsts } from "../shared/appconsts";
import * as moment from "moment";
import * as http from "./http";
import { MeekouResponse } from "./common.model";
export class MeekouApi {
  private baseUrl: string;
  /**
   *
   */
  constructor() {
    this.baseUrl = AppConsts.remoteServiceBaseUrl;
  }
  /**
   * @param body (optional)
   * @return Success
   */
  async loginByCode(body: LoginByCodeInput | undefined): Promise<LoginOutput> {
    let url_ = this.baseUrl + "/api/services/app/User/LoginByCode";
    url_ = url_.replace(/[?&]$/, "");
    const response = await http.post<MeekouResponse<LoginOutput>>(url_, body);
    if (response.parsedBody.success) {
      return response.parsedBody.result as LoginOutput;
    }
  }
}

export class LoginByCodeInput {
  inputCode: string | undefined;
  redirect: string | undefined;
}

export class LoginOutput {
  accessToken: string | undefined;
  user: UserDto;
  qrCodeOutput: QRCodeOutput;
}

export class UserDto {
  id: number;
  userName: string;
  name: string;
  surname: string;
  emailAddress: string;
  openId: string | undefined;
  isActive: boolean;
  fullName: string | undefined;
  lastLoginTime: moment.Moment | undefined;
  creationTime: moment.Moment;
  roleNames: string[] | undefined;
}

export class QRCodeOutput {
  qrCodeLink: string | undefined;
  content: string | undefined;
  outputCode: string | undefined;
}
