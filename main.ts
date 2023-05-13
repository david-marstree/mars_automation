import {
	App,
	Notice,
	Plugin,
	PluginSettingTab,
	Setting,
	TFolder,
} from "obsidian";

import { Graph } from "src/Graph";
// Remember to rename these classes and interfaces!

interface MARStreeAutomationSettings {
	plannerId: string;
}

const DEFAULT_SETTINGS: MARStreeAutomationSettings = {
	plannerId: "",
};

export default class MARStreeAutomation extends Plugin {
	settings: MARStreeAutomationSettings;

	async onload() {
		await this.loadSettings();

		const graph = new Graph();
		// add menu-item for file-menu
		this.registerEvent(
			this.app.workspace.on("file-menu", async (menu, file) => {
				const folderOrFile = this.app.vault.getAbstractFileByPath(
					file.path
				);
				if (folderOrFile instanceof TFolder) {
					menu.addItem((item) => {
						item.setIcon("folder")
							.setTitle("Tree: 生成每月維護發票")
							.onClick(async () => {
								graph
									.getPlanner(this.settings.plannerId)
									.then((lists) => {
										console.log("lists:", lists);
									});
							});
					});
				}
			})
		);

		// add setting tab
		this.addSettingTab(new MARStreeAutomationSettingTab(this.app, this));
	}

	onunload() {}

	async loadSettings() {
		this.settings = Object.assign(
			{},
			DEFAULT_SETTINGS,
			await this.loadData()
		);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}

class MARStreeAutomationSettingTab extends PluginSettingTab {
	plugin: MARStreeAutomation;

	constructor(app: App, plugin: MARStreeAutomation) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		containerEl.createEl("h2", { text: "Settings for my awesome plugin." });

		new Setting(containerEl)
			.setName("Planner ID")
			.setDesc("請輸入需要每月生成維護發票的Planner ID")
			.addText((text) =>
				text
					.setValue(this.plugin.settings.plannerId)
					.onChange(async (value) => {
						this.plugin.settings.plannerId = value;
						await this.plugin.saveSettings();
					})
			);
	}
}
