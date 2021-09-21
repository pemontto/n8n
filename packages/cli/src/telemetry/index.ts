/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
import TelemetryClient = require('@rudderstack/rudder-sdk-node');
import { IDataObject } from 'n8n-workflow';
import config = require('../../config');

export class Telemetry {
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	private client?: any;

	private instanceId: string;

	private pulseIntervalReference: NodeJS.Timeout;

	private executionCountsBuffer: IDataObject[] = [];

	constructor(instanceId: string) {
		this.instanceId = instanceId;

		const enabled = config.get('telemetry.enabled') as boolean;
		if (enabled) {
			this.client = new TelemetryClient(
				config.get('telemetry.config.backend.key') as string,
				config.get('telemetry.config.backend.url') as string,
			);

			this.pulseIntervalReference = setInterval(async () => {
				void this.pulse();
			}, 60 * 60 * 1000); // every hour
		}
	}

	private async pulse(): Promise<void> {
		if (!this.client) {
			return;
		}

		console.log('pulse started');

		await this.track('Workflow execution count', {
			count: this.executionCountsBuffer.length,
			executions: this.executionCountsBuffer,
		});
		this.executionCountsBuffer = [];

		console.log('pulse returning');
	}

	async trackWorkflowExecution(properties: IDataObject): Promise<void> {
		if (this.client) {
			this.executionCountsBuffer.push({ ...properties, timestamp: new Date().toISOString() });
		}
	}

	async trackN8nStop(): Promise<void> {
		clearInterval(this.pulseIntervalReference);
		await this.pulse();
		await this.track('User instance stopped');
	}

	async identify(traits?: IDataObject): Promise<void> {
		if (this.client) {
			await this.client.identify({
				userId: this.instanceId,
				anonymousId: '000000000000',
				traits,
			});
		}
	}

	async track(eventName: string, properties?: IDataObject): Promise<void> {
		if (this.client) {
			await this.client.track({
				userId: this.instanceId,
				event: eventName,
				anonymousId: '000000000000',
				properties,
			});
		}
	}
}
