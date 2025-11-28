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
