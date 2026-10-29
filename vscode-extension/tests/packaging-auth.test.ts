import { createRequire } from 'node:module';
import { afterEach, describe, expect, it, vi } from 'vitest';

// Exercise the CommonJS dependency instances used by vsce and Azure Identity.
const require = createRequire(import.meta.url);
const identity = require('@azure/identity') as typeof import('@azure/identity');
const identityRequire = createRequire(require.resolve('@azure/identity'));
const msal = identityRequire('@azure/msal-node') as typeof import('@azure/msal-node');
const base = Object.getPrototypeOf(msal.ConfidentialClientApplication.prototype) as {
	acquireTokenByCode: InstanceType<typeof msal.ConfidentialClientApplication>['acquireTokenByCode'];
	getAuthCodeUrl: InstanceType<typeof msal.ConfidentialClientApplication>['getAuthCodeUrl'];
};
const clientId = '00000000-0000-0000-0000-000000000001';
const tenantId = '00000000-0000-0000-0000-000000000002';
const scopes = ['https://management.azure.com/.default'];
const redirectUri = 'http://localhost:47111/callback';
const configuration = { auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}`, clientSecret: 'synthetic-secret' } };
const token = {
	authority: configuration.auth.authority,
	uniqueId: 'synthetic-user', tenantId, scopes,
	account: { homeAccountId: 'synthetic-home', localAccountId: 'synthetic-user', environment: 'login.microsoftonline.com', tenantId, username: 'synthetic@example.invalid' },
	idToken: '', idTokenClaims: {}, accessToken: 'synthetic-access-token', fromCache: false,
	expiresOn: new Date('2099-01-01T00:00:00Z'), tokenType: 'Bearer', correlationId: 'synthetic-correlation'
};

afterEach(() => vi.restoreAllMocks());

describe('packaging authentication dependency boundary', () => {
	it('does not forward attacker response routing or nonce metadata to code redemption', async () => {
		const redeem = vi.spyOn(base, 'acquireTokenByCode').mockResolvedValue(token);
		const client = new msal.ConfidentialClientApplication(configuration);
		const request = { scopes, redirectUri, code: 'request-code', nonce: 'request-nonce' };
		await client.acquireTokenByCode(request, {
			code: 'response-code', state: 'returned-state', nonce: 'response-nonce',
			cloud_instance_host_name: 'attacker.example.invalid', cloud_instance_name: 'attacker-cloud',
			cloud_graph_host_name: 'attacker-graph.example.invalid', msgraph_host: 'attacker-msgraph.example.invalid',
			client_info: 'attacker-client-info'
		});
		expect(redeem).toHaveBeenCalledExactlyOnceWith(request, {
			code: 'request-code', state: 'returned-state', nonce: 'request-nonce'
		});
	});

	it('rejects confidential-client instance-aware configuration', () => {
		const runtimeConfiguration = {
			...configuration, auth: { ...configuration.auth, instanceAware: true }
		};
		expect(() => new msal.ConfidentialClientApplication(runtimeConfiguration))
			.toThrow(expect.objectContaining({ errorCode: 'unsupported_cca_instance_aware' }));
	});

	it.each(['true', 'false', ''])('rejects instance_aware routing query value %j before authority discovery', async value => {
		const authorize = vi.spyOn(base, 'getAuthCodeUrl').mockResolvedValue('https://login.microsoftonline.com/synthetic');
		const client = new msal.ConfidentialClientApplication(configuration);
		await expect(client.getAuthCodeUrl({ scopes, redirectUri, extraQueryParameters: { instance_aware: value } }))
			.rejects.toMatchObject({ errorCode: 'unsupported_cca_instance_aware' });
		expect(authorize).not.toHaveBeenCalled();
	});

	it('preserves ordinary confidential code redemption and caller-provided code/nonce', async () => {
		const redeem = vi.spyOn(base, 'acquireTokenByCode').mockResolvedValue(token);
		const client = new msal.ConfidentialClientApplication(configuration);
		const request = { scopes, redirectUri, code: 'request-code', nonce: 'request-nonce' };
		expect(await client.acquireTokenByCode(request)).toBe(token);
		expect(redeem).toHaveBeenCalledOnce();
		expect(redeem.mock.calls[0][0]).toEqual(request);
		expect(redeem.mock.calls[0][1]).toBeUndefined();
	});

	it('preserves real returned-state validation before any network request', async () => {
		const network = {
			sendGetRequestAsync: vi.fn(async () => { throw new Error('Unexpected authentication network request'); }),
			sendPostRequestAsync: vi.fn(async () => { throw new Error('Unexpected authentication network request'); })
		};
		const client = new msal.ConfidentialClientApplication({ ...configuration, system: { networkClient: network } });
		await expect(client.acquireTokenByCode({ scopes, redirectUri, code: 'request-code', state: 'expected-state' }, {
			code: 'response-code', state: 'wrong-state', cloud_instance_host_name: 'attacker.example.invalid'
		})).rejects.toMatchObject({ errorCode: 'state_mismatch' });
		expect(network.sendGetRequestAsync).not.toHaveBeenCalled();
		expect(network.sendPostRequestAsync).not.toHaveBeenCalled();
	});

	it.each([false, true])('preserves Azure Identity authorization-code adaptation with confidential=%j', async confidential => {
		const prototype = confidential ? msal.ConfidentialClientApplication.prototype : msal.PublicClientApplication.prototype;
		const redeem = vi.spyOn(prototype, 'acquireTokenByCode').mockResolvedValue(token);
		const credential = confidential
			? new identity.AuthorizationCodeCredential(tenantId, clientId, 'synthetic-secret', 'request-code', redirectUri)
			: new identity.AuthorizationCodeCredential(tenantId, clientId, 'request-code', redirectUri);
		expect(await credential.getToken(scopes)).toMatchObject({ token: token.accessToken, expiresOnTimestamp: token.expiresOn.getTime() });
		expect(redeem).toHaveBeenCalledOnce();
		expect(redeem.mock.calls[0][0]).toMatchObject({ code: 'request-code', redirectUri, scopes });
		expect(redeem.mock.calls[0]).toHaveLength(1);
	});

	it('preserves Azure Identity client-secret token adaptation', async () => {
		const acquire = vi.spyOn(msal.ConfidentialClientApplication.prototype, 'acquireTokenByClientCredential').mockResolvedValue(token);
		const credential = new identity.ClientSecretCredential(tenantId, clientId, 'synthetic-secret');
		expect(await credential.getToken(scopes)).toMatchObject({ token: token.accessToken, expiresOnTimestamp: token.expiresOn.getTime() });
		expect(acquire).toHaveBeenCalledOnce();
		expect(acquire.mock.calls[0][0]).toMatchObject({ scopes });
	});

	it('preserves Azure Identity interactive requests without removed responseMode', async () => {
		const acquire = vi.spyOn(msal.PublicClientApplication.prototype, 'acquireTokenInteractive').mockResolvedValue(token);
		const credential = new identity.InteractiveBrowserCredential({ tenantId, clientId, redirectUri });
		expect(await credential.getToken(scopes)).toMatchObject({ token: token.accessToken, expiresOnTimestamp: token.expiresOn.getTime() });
		expect(acquire).toHaveBeenCalledOnce();
		expect(acquire.mock.calls[0][0]).not.toHaveProperty('responseMode');
	});
});
