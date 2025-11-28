# Repo Steps for SPFx Marketplace integration - TrackingID#2511140040006028

Sample SPFx project available here: https://github.com/allenj3304/Marketplace-Test-spfx

## Part 1: Steps to create SPFx web part that calls Microsoft Graph usageRights API
1. Setup SPFx dev env: (https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
2. Create a new SPFx web part project: (https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part)
	Using PowerShell or VS Code Terminal
	a. Create new solution directory and change current directory to the new folder> md Marketplace-Test-SPFx  >cd Marketplace-test-SPFx
	b. Create a new project by running the Yeoman SharePoint Generator> yo @microsoft/sharepoint
		Select options:
		* Name: use default
		* Type: WebPart
		* Web part name: use default (HelloWorld)
		* Framework: React
	c. Replace '{tenantDomain}' in ./.vscode/launch.json and ./config/serve.json with you testing tenant. e.g. testdomain.sharepoint.com/sites/mytest
3. Access Microsoft Graph: (https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
	a. Open ./config/package-solution.json
	b. Update the solution section to include permission grant request as below
``` json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    //...

    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "Organization.Read.All"
      }
    ],


    //...
  }
}
```

4. Update code to call MS Graph usageRights
	a. Pass context to HellowWorld.tsx: Add 'context: WebPartContext;' property to IHelloWorldProps.ts, Open HelloWorldWebPart.ts and add 'context: this.context' to  React.createElement, Open HelloWorld.tsx and add context to deconstruction of this.props.
	b. Open HelloWorld.tsx and replace with following code:
``` typescript
import * as React from 'react';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps } from './IHelloWorldProps';
import { PrimaryButton } from '@fluentui/react';

interface IUsageRight {
    id: string;
    skuId: string;
    skuPartNumber: string;
    catalogId: string;
    serviceIdentifier: string;
    servicePlanId: string;
    state: string;
    assignedDateTime: string;
}

interface IUsageRightsResponse {
    value: IUsageRight[];
}

export default class HelloWorld extends React.Component<IHelloWorldProps> {

  private usageRights: IUsageRightsResponse | undefined;

  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      hasTeamsContext,
      context
    } = this.props;

    async function getUsageRights(): Promise<IUsageRightsResponse | undefined> {
      try {
        const graphClient = await context.msGraphClientFactory.getClient("3");
        const result = await graphClient
          .api(`/me/usageRights`)
          .version('beta')
          .get();

        return result.value as IUsageRightsResponse;

      } catch (error) {
        console.error("Exception while fetching organization data:", error);
      }
    };

    function handleButton(event: React.MouseEvent<HTMLButtonElement>): void {
      event.preventDefault();
      this.usageRights = getUsageRights();
    }

    return (
      <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <PrimaryButton text="Get Usage Rights" onClick={handleButton} />

          {this.usageRights && this.usageRights.value.length > 0
            ? (
                <div>
                  <h3>Usage Rights:</h3>
                  <ul>
                    {this.usageRights.value.map((right, index) => (
                      <li key={index}>{right.serviceIdentifier} - {right.state}</li>
                    ))}
                  </ul>
                </div>
            )
            : (<p>No usage rights found.</p>)}
        </div>
      </section>
    );
  }
}
```
5. Package and deploy the solution to your app catalog: (https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page)
  a. In the terminal run 'gulp bundle --ship' then 'gulp package-solution --ship'
  b. Upload the generated .sppkg file from the ./sharepoint/solution folder to your tenant app catalog site.
  c. After add ing to app catalog follow steps to Approve the API permissions.
  d. Add the app to a SharePoint site and test.

## Part 2: Marketplace offer purchase.
Purchase a SaaS offer with free plan. You can use my test offer. <b>Send me the test tenant ID and I will add it to the preview Audience.</b>
https://marketplace.microsoft.com/en-us/product/saas/xrsolutions.xrs-acronym-analysis-test-preview?tab=PlansAndPrice&flightCodes=3624c142-fd12-4fe9-bf20-28483caa4ed6
Alternatively, create your own test offer following these steps:
1. Create a new Marketplace offer in Partner Center with the following offer setup:
  a. Offer type: SaaS
  b. Select 'Yes, I would like to sell through Microsoft and have Microsoft host transactions on my behalf'
  c. Select 'Yes, I would like Microsoft to manage customer licenses on my behalf'
  d. Add a new plan with the following plan setup:
    i. Plan pricing: Per User
    ii. Price per charge: 0 USD
2. Use Graph Explorer to verify the usageRights is updated after purchase.
  a. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
  b. Sign in with the user that purchased the offer.
  c. Run the following query: GET https://graph.microsoft.com/beta/me/usageRights
  d. Verify the response contains an entry for the purchased plan.
3. Go back to the SPFx web part and click the 'Get Usage Rights' button to test usageRights in the web part.


