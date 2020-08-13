import "isomorphic-fetch";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

/**
  * @module TokenPassThroughAuthProvider
  */

export class TokenPassThroughAuthProvider implements AuthenticationProvider {

    private authtoken: string;

    /**
	 * @public
	 * @constructor
	 * Creates an instance of TokenPassThroughAuthProvider
	 * @param {string} authtoken - A valid and appropriately scoped Graph API auth token
	 * @returns An instance of CustomAuthenticationProvider
	 */
	public constructor(authtoken: string) {
		this.authtoken = authtoken;
    }
    
	/**
	 * @public
	 * @async
	 * To get the access token
	 * @returns The promise that resolves to an access token
	 */
	public getAccessToken(): Promise<any> {
		return new Promise((resolve: (accessToken: string) => void, reject: (error: any) => void) => {
            if (this.authtoken) {
                resolve(this.authtoken);
            } else {
                reject("No Auth Token Provided");
            }
		});
	}

}