import * as msal from "@azure/msal-node";
import * as msalCommon from "@azure/msal-common";
import { Client } from "@microsoft/microsoft-graph-client";
import {
	TodoTask,
	TodoTaskList,
	PlannerTask,
} from "@microsoft/microsoft-graph-types";

import { DataAdapter, Notice } from "obsidian";
import { MicrosoftAuthModal } from "./MicrosoftAuthModal";

export class Graph {
	private client: Client;

	constructor() {
		new MicrosoftClientProvider().getClient().then((client) => {
			this.client = client;
		});
	}

	// List operation
	async getLists(
		searchPattern?: string
	): Promise<TodoTaskList[] | undefined> {
		const endpoint = "/me/todo/lists";
		const todoLists = (await this.client.api(endpoint).get())
			.value as TodoTaskList[];
		return await Promise.all(
			todoLists.map(async (taskList) => {
				const containedTasks = await this.getListTasks(
					taskList.id,
					searchPattern
				);
				return {
					...taskList,
					tasks: containedTasks,
				};
			})
		);
	}

	// Task operation
	async getListTasks(
		listId: string | undefined,
		searchText?: string
	): Promise<TodoTask[] | undefined> {
		if (!listId) return;
		const endpoint = `/me/todo/lists/${listId}/tasks`;
		if (!searchText) return;
		const res = await this.client
			.api(endpoint)
			.filter(searchText)
			.get()
			.catch((err) => {
				new Notice("获取失败，请检查同步列表是否已删除");
				return;
			});
		if (!res) return;
		return res.value as TodoTask[];
	}

	// Get Planner
	async getPlanner(plannerId: string): Promise<PlannerTask[] | undefined> {
		if (!plannerId) return;
		const endpoint = `/planner/plans/${plannerId}/tasks`;
		const res = await this.client
			.api(endpoint)
			.get()
			.catch((err) => {
				new Notice("獲取失敗，請檢查同步Planner 是否已刪除");
				return;
			});
		if (!res) return;
		return res.value as PlannerTask[];
	}
}

export class MicrosoftClientProvider {
	private readonly clientId = "d782e01f-232d-46cc-9904-5ee4503f01c9";
	private readonly authority =
		"https://login.microsoftonline.com/2d34ea4c-f8d5-4ff8-b338-c2a1420325fd";
	private readonly clientSecret = "MrX8Q~NXo-ey34rQHFwDSBFo5QGmP.4ztwOsKb-J";
	private readonly scopes: string[] = [
		"Group.Read.All",
		"Group.ReadWrite.All",
		"Tasks.Read",
		"Tasks.ReadWrite",
		"openid",
		"profile",
	];
	private readonly pca: msal.PublicClientApplication;
	private readonly adapter: DataAdapter;
	private readonly cachePath: string;

	constructor() {
		this.adapter = app.vault.adapter;
		console.log("adapter", app.vault.configDir);
		this.cachePath = `${app.vault.configDir}/Microsoft_cache.json`;

		const beforeCacheAccess = async (
			cacheContext: msalCommon.TokenCacheContext
		) => {
			if (await this.adapter.exists(this.cachePath)) {
				cacheContext.tokenCache.deserialize(
					await this.adapter.read(this.cachePath)
				);
			}
		};
		const afterCacheAccess = async (
			cacheContext: msalCommon.TokenCacheContext
		) => {
			if (cacheContext.cacheHasChanged) {
				await this.adapter.write(
					this.cachePath,
					cacheContext.tokenCache.serialize()
				);
			}
		};
		const cachePlugin = {
			beforeCacheAccess,
			afterCacheAccess,
		};
		const config = {
			auth: {
				clientId: this.clientId,
				authority: this.authority,
				clientSecret: this.clientSecret,
			},
			cache: {
				cachePlugin,
			},
		};
		this.pca = new msal.PublicClientApplication(config);
	}

	private async getAccessToken() {
		const msalCacheManager = this.pca.getTokenCache();
		if (await this.adapter.exists(this.cachePath)) {
			msalCacheManager.deserialize(
				await this.adapter.read(this.cachePath)
			);
		}
		const accounts = await msalCacheManager.getAllAccounts();
		if (accounts.length == 0) {
			return await this.authByDevice();
		} else {
			return await this.authByCache(accounts[0]);
		}
	}
	private async authByDevice(): Promise<string> {
		const deviceCodeRequest = {
			deviceCodeCallback: (response: msalCommon.DeviceCodeResponse) => {
				new Notice("设备代码已复制到剪贴板,请在打开的浏览器界面输入");
				navigator.clipboard.writeText(response["userCode"]);
				new MicrosoftAuthModal(
					response["userCode"],
					response["verificationUri"]
				).open();
				console.log("设备代码已复制到剪贴板", response["userCode"]);
			},
			scopes: this.scopes,
		};
		return await this.pca
			.acquireTokenByDeviceCode(deviceCodeRequest)
			.then((res) => {
				return res == null ? "error" : res["accessToken"];
			});
	}

	private async authByCache(account: msal.AccountInfo): Promise<string> {
		const silentRequest = {
			account: account,
			scopes: this.scopes,
		};
		return await this.pca
			.acquireTokenSilent(silentRequest)
			.then((res) => {
				return res == null ? "error" : res["accessToken"];
			})
			.catch(async (err) => {
				return await this.authByDevice();
			});
	}

	public async getClient() {
		const authProvider = async (
			callback: (arg0: string, arg1: string) => void
		) => {
			const accessToken = await this.getAccessToken();
			const error = " ";
			callback(error, accessToken);
		};
		return Client.init({
			authProvider,
		});
	}
}
