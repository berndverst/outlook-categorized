import "isomorphic-fetch";
import { Client, PageCollection, PageIterator, PageIteratorCallback } from "@microsoft/microsoft-graph-client";
import { TokenPassThroughAuthProvider } from "./TokenPassThroughAuthProvider";
import ReactDOM from 'react-dom';
import React from 'react';

export class OutlookCategorizer {

    private client: Client;
    private categories: Set<String>
    private folders: Set<String>
    private existingCategoryFolders: Set<String>
    private createdCategoryFolders: Set<String>

    public constructor(authtoken: string) {
        var auth = new TokenPassThroughAuthProvider(authtoken);
        this.client = Client.initWithMiddleware({ authProvider: auth });
        this.categories = new Set();
        this.folders = new Set();
        this.existingCategoryFolders = new Set();
        this.createdCategoryFolders = new Set();

        this.getUser()
        this.createMailSearchFolders();
    }

    public async getUser(): Promise<void> {
        try {
            let userDetails = await this.client.api('/me').get()
            console.log(userDetails);
        } catch (error) {
            throw error;
        }
    }

    public async getCategories(): Promise<void> {
        this.categories = new Set();

        try {
            // Makes request to fetch mails list. Which is expected to have multiple pages of data.
            let response: PageCollection = await this.client.api("/me/outlook/masterCategories?$top=500").get();
            // A callback function to be called for every item in the collection. This call back should return boolean indicating whether not to continue the iteration process.
            let callback: PageIteratorCallback = (data) => {
                this.categories.add(data.displayName);
                return true;
            };
            // Creating a new page iterator instance with client a graph client instance, page collection response from request and callback
            let pageIterator = new PageIterator(this.client, response, callback);
            // This iterates the collection until the nextLink is drained out.
            pageIterator.iterate();
        } catch (e) {
            throw e;
        }
    }

    public async getFolders(): Promise<void> {
        this.folders = new Set();

        try {
            // Makes request to fetch mails list. Which is expected to have multiple pages of data.
            let response: PageCollection = await this.client.api("/me/mailFolders?$top=500").get();
            // A callback function to be called for every item in the collection. This call back should return boolean indicating whether not to continue the iteration process.
            let callback: PageIteratorCallback = (data) => {
                this.folders.add(data.displayName.toLowerCase())
                return true;
            };
            // Creating a new page iterator instance with client a graph client instance, page collection response from request and callback
            let pageIterator = new PageIterator(this.client, response, callback);
            // This iterates the collection until the nextLink is drained out.
            pageIterator.iterate();

            if (!pageIterator.isComplete()) {
                pageIterator.resume();
            }
        } catch (e) {
            throw e;
        }
    }

    public async getRootFolderId(): Promise<String> {
        try {
            let inbox = await this.client.api('/me/mailFolders/inbox').get()
            return inbox.parentFolderId
        } catch (e) {
            throw e;
        }
    }

    public async createMailSearchFolders(): Promise<void> {
        await this.getCategories();
        await this.getFolders();
        let rootFolderId = await this.getRootFolderId();
        let categories = Array.from(this.categories.values());
        let folders = Array.from(this.folders.values());

        this.existingCategoryFolders = new Set();
        this.createdCategoryFolders = new Set();


        for (let category of categories) {
            if (folders.indexOf(category.toLowerCase()) > -1) {
                console.log(category + " folder already exists! Skipping this category.");
                this.existingCategoryFolders.add(category);
            }
            else {
                var searchFolderCreateRequest = {
                    "@odata.type": "microsoft.graph.mailSearchFolder",
                    "displayName": category,
                    "includeNestedFolders": true,
                    "sourceFolderIds": [rootFolderId],
                    "filterQuery": `categories/any(t: t eq '${category}')`
                }

                console.log(searchFolderCreateRequest);

                try {
                    await this.client.api(`/me/mailFolders/${rootFolderId}/childFolders`).post(searchFolderCreateRequest);
                    this.createdCategoryFolders.add(category);
                }
                catch (e) {
                    throw e;
                }

            }

        }

        ReactDOM.render(
            <React.StrictMode>
                <p><strong>Existing virtual category folders or conflicting folder names:</strong></p>
                <p>
                    {Array.from(this.existingCategoryFolders.values()).join(', ')}
                </p>
                <p><strong>Newly created virtual mail folders:</strong></p>
                <p>
                    {Array.from(this.createdCategoryFolders.values()).join(', ')}

                </p>
            </React.StrictMode>,
            document.getElementById('results'));
    }


}