figma.showUI(__html__, { width: 500, height: 500 });
let DEBUG_BASIC = true; // Set to false to disable basic logging
let DEBUG_DETAILED = false; // Set to false to disable detailed logging

function setDebugLevel(level: string) {
	DEBUG_BASIC = level === "basic" || level === "detailed";
	DEBUG_DETAILED = level === "detailed";
}

// Add an event listener to listen for messages from the HTML UI

figma.ui.onmessage = async (msg: {
	isDebugEnabled: any;
	type: string;
	mode?: string;
	data?: any;
}) => {
	debugLogDetailed("Received message:", msg);
	// Handle the message based on the type
	debugLogDetailed("Message type received:", msg.type);
	switch (msg.type) {
		case "sync-kpi":
			debugLogDetailed("Handling sync-kpi message");
			debugLogDetailed("Mode:", msg.mode);
			debugLogDetailed("Data:", msg.data);
			const sheetName = msg.data?.sheetName;
			if (!sheetName) {
				console.error("Sheet name is missing in the data");
				return;
			}
			debugLogDetailed("Sheet Name:", sheetName);
			const collection = getOrCreateCollection(sheetName);
			debugLogDetailed("Collection:", collection);
			if (!collection) {
				console.error("Failed to get or create collection");
				return;
			}
			const kpiType = getKpiType(sheetName);
			if (!kpiType) {
				console.error("Failed to determine KPI type");
				return;
			}
			if (kpiType === "BM") {
				if (isTextOnlyKPI(sheetName)) {
					debugLogDetailed("Text only KPI sheet:", sheetName);
					debugLogDetailed("Data:", msg.data?.records);
					createBMTextKPIVariablesFromData(msg.data?.records, collection);
					cleanUpCollection(sheetName);
				} else {
					debugLogDetailed("BM KPI sheet:", sheetName);
					debugLogDetailed("Data:", msg.data?.records);
					createBMVariablesFromData(msg.data?.records, collection, kpiType);
					cleanUpCollection(sheetName);
				}
			} else if (kpiType === "HC") {
				if (sheetName === "ExperienceHealthCheck") {
					debugLogDetailed("ExperienceHealthCheck sheet:", sheetName);
					debugLogDetailed("Data:", msg.data?.records);
					createExperienceHealthCheckVariablesFromData(
						msg.data?.records,
						collection
					);
					cleanUpCollection(sheetName);
				} else {
					debugLogDetailed("HC KPI sheet:", sheetName);
					debugLogDetailed("Data:", msg.data?.records);
					createHCVariablesFromData(msg.data?.records, collection);
					cleanUpCollection(sheetName);
				}
			} else if (kpiType === "IN") {
				debugLogDetailed("Initiatives sheet:", sheetName);
				debugLogDetailed("Data:", msg.data?.records);
				createInitativesVariablesFromData(msg.data?.records, collection);
				cleanUpCollection(sheetName);
			} else {
				console.error("Invalid KPI type");
			}
			break;

		case "BM":
			debugLogDetailed("Handling Business Metrics message");
			debugLogDetailed("Mode:", msg.mode);
			debugLogDetailed("Data:", msg.data);
			debugLogDetailed("Number of sheets:", Object.keys(msg.data).length);
			const data = msg.data;
			Object.keys(data).forEach((sheet: string) => {
				const collection = getOrCreateCollection(sheet);
				debugLogDetailed("Collection:", collection);
				if (!collection) {
					console.error("Failed to get or create collection");
					return;
				} else {
					const kpiType = getKpiType(sheet);
					if (!kpiType) {
						console.error("Failed to determine KPI type for :", sheet);
						return;
					}
					if (kpiType === "BM") {
						if (isTextOnlyKPI(sheet)) {
							debugLogDetailed("Text only KPI sheet:", sheet);
							createBMTextKPIVariablesFromData(data[sheet], collection);
							cleanUpCollection(sheet);
						} else {
							createBMVariablesFromData(data[sheet], collection, kpiType);
							cleanUpCollection(sheet);
						}
					} else {
						debugLogDetailed(
							"Sheet name:",
							sheet + " is not a Business Metrics sheet"
						);
					}
				}
			});
			break;

		case "HC":
			debugLogDetailed("Handling Health Check message");
			debugLogDetailed("Mode:", msg.mode);
			debugLogDetailed("Data:", msg.data);
			debugLogDetailed("Number of sheets:", Object.keys(msg.data).length);
			const hcData = msg.data;
			Object.keys(hcData).forEach((sheet: string) => {
				const collection = getOrCreateCollection(sheet);
				debugLogDetailed("Collection:", collection);
				if (!collection) {
					console.error("Failed to get or create collection");
					return;
				} else {
					const kpiType = getKpiType(sheet);
					if (!kpiType) {
						console.error("Failed to determine KPI type for :", sheet);
						return;
					}
					if (kpiType === "HC") {
						if (sheet === "ExperienceHealthCheck") {
							debugLogDetailed("ExperienceHealthCheck sheet:", sheet);
							createExperienceHealthCheckVariablesFromData(
								hcData[sheet],
								collection
							);
							cleanUpCollection(sheet);
						} else {
							createHCVariablesFromData(hcData[sheet], collection);
							cleanUpCollection(sheet);
						}
					} else {
						debugLogDetailed(
							"Sheet name:",
							sheet + " is not a Health Check sheet"
						);
					}
				}
			});
			break;
		case "IN":
			debugLogDetailed("Handling Initiatives message");
			debugLogDetailed("Mode:", msg.mode);
			debugLogDetailed("Data:", msg.data);
			debugLogDetailed("Number of sheets:", Object.keys(msg.data).length);
			const inData = msg.data;
			Object.keys(inData).forEach((sheet: string) => {
				const collection = getOrCreateCollection(sheet);
				debugLogDetailed("Collection:", collection);
				if (!collection) {
					console.error("Failed to get or create collection");
					return;
				} else {
					const kpiType = getKpiType(sheet);
					if (!kpiType) {
						console.error("Failed to determine KPI type for :", sheet);
						return;
					}
					if (kpiType === "IN") {
						createInitativesVariablesFromData(inData[sheet], collection);
						cleanUpCollection(sheet);
					} else {
						debugLogDetailed(
							"Sheet name:",
							sheet + " is not an Initiatives sheet"
						);
					}
				}
			});
			break;

		case "toggle-debug":
			const isDebugEnabled = msg.isDebugEnabled;
			setDebugLevel(isDebugEnabled ? "detailed" : "basic");
			debugLogBasic(
				`Debug mode set to ${isDebugEnabled ? "detailed" : "basic"}`
			);
			break;
		default:
			console.error("Received unknown message type:", msg.type);
	}
};

// Functions to create variables from data based on different KPI types

// Function to create Initiatives variables from data

function createInitativesVariablesFromData(
	data: any[],
	collection: VariableCollection
) {
	if (!data || !Array.isArray(data)) {
		console.error("Invalid data format");
		return;
	}
	let totalRecordsProcessed = 0;
	data.forEach((record) => {
		totalRecordsProcessed++;
		const mode = getOrAddModeForCollection(collection, "INVALUE");
		var variableName = "";
		let initiativeName = record.Initiative;
		let metricName = record.Metric;
		let graphType = record.GTYPE;
		Object.entries(record).forEach(([key, value]) => {
			switch (key) {
				case "Metric":
				case "Description":
				case "FY23":
				case "FY24_YTD":
				case "Target":
				case "Target_Percentage":
					variableName = initiativeName + "/" + metricName + "/" + key;
					updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "Progress_Bar":
					variableName = initiativeName + "/" + metricName + "/" + key;
					updateOrCreateVariable(
						variableName,
						Number(value),
						mode,
						"FLOAT",
						collection
					);
					break;
				case "Graph/Nov_23":
				case "Graph/Dec_23":
				case "Graph/Jan_24":
				case "Graph/Feb_24":
				case "Graph/Mar_24":
				case "Graph/Apr_24":
				case "Graph/May_24":
				case "Graph/Jun_24":
				case "Graph/Jul_24":
				case "Graph/Aug_24":
				case "Graph/Sep_24":
				case "Graph/Oct_24":
					if (graphType === "M") {
						variableName = initiativeName + "/" + metricName + "/" + key;
						updateOrCreateVariable(
							variableName,
							String(value),
							mode,
							"STRING",
							collection
						);
					}
					break;
				case "Graph/Quarter_1":
				case "Graph/Quarter_2":
				case "Graph/Quarter_3":
				case "Graph/Quarter_4":
					if (graphType === "Q") {
						variableName = initiativeName + "/" + metricName + "/" + key;
						updateOrCreateVariable(
							variableName,
							String(value),
							mode,
							"STRING",
							collection
						);
					}
					break;
				case "GSize/Nov_23":
				case "GSize/Dec_23":
				case "GSize/Jan_24":
				case "GSize/Feb_24":
				case "GSize/Mar_24":
				case "GSize/Apr_24":
				case "GSize/May_24":
				case "GSize/Jun_24":
				case "GSize/Jul_24":
				case "GSize/Aug_24":
				case "GSize/Sep_24":
				case "GSize/Oct_24":
					if (graphType === "M") {
						variableName = initiativeName + "/" + metricName + "/" + key;
						updateOrCreateVariable(
							variableName,
							Number(value),
							mode,
							"FLOAT",
							collection
						);
					}
					break;
				case "GSize/Quarter_1":
				case "GSize/Quarter_2":
				case "GSize/Quarter_3":
				case "GSize/Quarter_4":
					if (graphType === "Q") {
						variableName = initiativeName + "/" + metricName + "/" + key;
						updateOrCreateVariable(
							variableName,
							Number(value),
							mode,
							"FLOAT",
							collection
						);
					}
					break;
				default:
					debugLogDetailed("Invalid key:", key);
					break;
			}
		});
	});
	debugLogBasic(
		`Processed KPI: ${collection.name}, Total records: ${totalRecordsProcessed}`
	);
}

// Function to create BM variables from data
function createBMVariablesFromData(
	data: any[],
	collection: VariableCollection,
	kpiType: string
) {
	if (!data || !Array.isArray(data)) {
		console.error("Invalid data format");
		return;
	}
	let totalRecordsProcessed = 0;
	let modifiedRecordsCount = 0;
	data.forEach((record) => {
		totalRecordsProcessed++;
		const mode = getOrAddModeForCollection(collection, record.Month);
		const yearValue = record.Year;
		var variableName = "";
		var isModified = false;
		let isRecordModified = false;
		debugLogDetailed("Mode Vaue: ", mode);
		Object.entries(record).forEach(([key, value]) => {
			switch (key) {
				case "Formatted Value":
					variableName = "Value/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "MOM":
				case "YOY":
					variableName = `${key}/Value/${yearValue}`;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "MOM Trend":
				case "YOY Trend":
					variableName = `${
						key === "MOM Trend" ? "MOM" : "YOY"
					}/Trend/${yearValue}`;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
					variableName = "YOY/Trend/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "M1":
				case "M2":
				case "M3":
				case "M4":
				case "M5":
				case "M6":
					variableName = key + "/Value/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "M1-Label":
				case "M2-Label":
				case "M3-Label":
				case "M4-Label":
				case "M5-Label":
				case "M6-Label":
					variableName = key.replace("-Label", "") + "/MonthLabel/" + yearValue;
					updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "GM1":
				case "GM2":
				case "GM3":
				case "GM4":
				case "GM5":
				case "GM6":
					variableName = key.replace("GM", "M") + "/GValue/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						Number(value),
						mode,
						"FLOAT",
						collection
					);
					break;
				default:
					debugLogDetailed("Invalid key:", key);
					break;
			}
			if (isModified) {
				isRecordModified = true;
			}
		});
		if (isRecordModified) {
			modifiedRecordsCount++;
		}
	});
	debugLogBasic(
		`Processed KPI: ${collection.name}, Total records: ${totalRecordsProcessed}, Modified records: ${modifiedRecordsCount}`
	);
}

// Function to create BM text KPI variables from data
function createBMTextKPIVariablesFromData(
	data: any[],
	collection: VariableCollection
) {
	if (!data || !Array.isArray(data)) {
		console.error("Invalid data format");
		return;
	}
	let totalRecordsProcessed = 0;
	let modifiedRecordsCount = 0;
	data.forEach((record) => {
		totalRecordsProcessed++;
		const mode = getOrAddModeForCollection(collection, record.Month);
		const yearValue = record.Year;
		var variableName = "";
		let isRecordModified = false;

		Object.entries(record).forEach(([key, value]) => {
			if (key === "Month" || key === "Year") {
				return;
			}
			variableName = key + "/" + yearValue;
			debugLogDetailed("Variable Name: ", variableName);
			const isModified = updateOrCreateVariable(
				variableName,
				String(value),
				mode,
				"STRING",
				collection
			);
			if (isModified) {
				isRecordModified = true;
			}
		});
		if (isRecordModified) {
			modifiedRecordsCount++;
		}
	});
	debugLogBasic(
		`Processed KPI: ${collection.name}, Total records: ${totalRecordsProcessed}, Modified records: ${modifiedRecordsCount}`
	);
}

// Function to create HC variables from data
function createHCVariablesFromData(
	data: any[],
	collection: VariableCollection
) {
	if (!data || !Array.isArray(data)) {
		console.error("Invalid data format");
		return;
	}
	let totalRecordsProcessed = 0;
	let modifiedRecordsCount = 0;
	data.forEach((record) => {
		totalRecordsProcessed++;
		const mode = getOrAddModeForCollection(collection, record.Month);
		const yearValue = record.Year;
		var variableName = "";
		let isModified = false;
		let isRecordModified = false;

		Object.entries(record).forEach(([key, value]) => {
			switch (key) {
				case "EFT Create Payment":
				case "Create EFT Template":
				case "Manage EFT Template":
				case "Interac E-Transfer":
				case "Single Wire":
				case "Account Transfer":
					variableName = key + "/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						String(value),
						mode,
						"STRING",
						collection
					);
					break;
				case "EFT Create Payment/GValue":
				case "Create EFT Template/GValue":
				case "Manage EFT Template/GValue":
				case "Interac E-Transfer/GValue":
				case "Single Wire/GValue":
				case "Account Transfer/GValue":
					variableName = key + "/" + yearValue;
					isModified = updateOrCreateVariable(
						variableName,
						Number(value),
						mode,
						"FLOAT",
						collection
					);
					break;
				default:
					debugLogDetailed("Invalid key:", key);
					break;
			}
			if (isModified) {
				isRecordModified = true;
			}
		});
		if (isRecordModified) {
			modifiedRecordsCount++;
		}
	});
	debugLogBasic(
		`Processed KPI: ${collection.name}, Total records: ${totalRecordsProcessed}, Modified records: ${modifiedRecordsCount}`
	);
}

// Function to create Experience Health Check variables from data
function createExperienceHealthCheckVariablesFromData(
	data: any[],
	collection: VariableCollection
) {
	if (!data || !Array.isArray(data)) {
		console.error("Invalid data format");
		return;
	}
	let totalRecordsProcessed = 0;
	let modifiedRecordsCount = 0;

	data.forEach((record) => {
		totalRecordsProcessed++;
		const mode = getOrAddModeForCollection(collection, "HCVALUE");
		var variableName = "";
		var title = record.Title + "/";
		let isRecordModified = false;

		Object.entries(record).forEach(([key, value]) => {
			if (key === "Title") {
				return;
			}
			variableName = title + key;
			debugLogDetailed("Variable Name: ", variableName);
			const isModified = updateOrCreateVariable(
				variableName,
				String(value),
				mode,
				"STRING",
				collection
			);
			if (isModified) {
				isRecordModified = true;
			}
		});
		if (isRecordModified) {
			modifiedRecordsCount++;
		}
	});
	debugLogBasic(
		`Processed KPI: ${collection.name}, Total records: ${totalRecordsProcessed}, Modified records: ${modifiedRecordsCount}`
	);
}

// Function to clean up the collection by removing modes that are not required
function cleanUpCollection(sheetName: any) {
	const collections = figma.variables.getLocalVariableCollections();
	let collection = collections.find((col) => col.name === sheetName);
	if (collection) {
		getAllModesForCollection(collection).forEach((mode) => {
			if (
				![
					"JAN",
					"FEB",
					"MAR",
					"APR",
					"MAY",
					"JUN",
					"JUL",
					"AUG",
					"SEP",
					"OCT",
					"NOV",
					"DEC",
					"HCVALUE",
					"INVALUE",
				].includes(mode.name)
			) {
				collection.removeMode(mode.id);
			}
		});
	}
}

// Only supporting functions below this line

function isTextOnlyKPI(sheet: string): boolean {
	const textOnlyKPIs = ["SelfServe", "ServiceNowCategories", "Comments"];
	return textOnlyKPIs.includes(sheet);
}
// Function to get KPI type based on sheet name
function getKpiType(sheetName: string): string | undefined {
	const kpiGroups: { [key: string]: string[] } = {
		BM: [
			"HDCalls",
			"SNTickets",
			"SelfServe",
			"ServiceNowCategories",
			"Comments",
		],
		HC: ["TaskTime", "CompletionRate", "ExperienceHealthCheck"],
		IN: ["Initiatives"],
	};

	for (const kpiType in kpiGroups) {
		if (kpiGroups.hasOwnProperty(kpiType)) {
			const sheets = kpiGroups[kpiType];
			if (sheets.includes(sheetName)) {
				return kpiType;
			}
		}
	}

	return undefined;
}

// Function to get or create a collection
function getOrCreateCollection(sheetName: string) {
	const collections = figma.variables.getLocalVariableCollections();
	let collection = collections.find((col) => col.name === sheetName);

	if (!collection) {
		debugLogDetailed(
			`No collection found with the name: ${sheetName}. Creating a new collection.`
		);
		collection = figma.variables.createVariableCollection(sheetName);
		debugLogDetailed(`Collection created: ${collection.name}`);
	} else {
		debugLogDetailed(`Collection found: ${collection.name}`);
	}

	return collection;
}

function getAllModesForCollection(collection: VariableCollection) {
	return collection.modes.map((mode: { modeId: string; name: string }) => ({
		name: mode.name,
		id: mode.modeId,
	}));
}

function getOrAddModeForCollection(
	collection: VariableCollection,
	modeName: string
) {
	const mode = collection.modes.find((mode) => mode.name === modeName);
	if (mode) {
		return mode.modeId;
	} else {
		return collection.addMode(modeName);
	}
}

// Function to update or create a variable with mode consideration
function updateOrCreateVariable(
	variableName: string,
	newValue: any,
	mode: string,
	valueType: string,
	collection: VariableCollection
): boolean {
	debugLogDetailed(`Updating variable: ${variableName} in mode: ${mode}`);
	var variables = figma.variables.getLocalVariables();
	const filteredVariables = variables.filter(
		(variable) => variable.variableCollectionId === collection.id
	);
	const existingVariable = filteredVariables.find(
		(variable) => variable.name === variableName
	);

	if (existingVariable) {
		// debugLogDetailed(`Variable found: ${existingVariable}`);
		const isModeAvailable = existingVariable.valuesByMode[mode];
		debugLogDetailed(`Mode available: ${isModeAvailable}`);
		if (isNotEmpty(isModeAvailable) && valueType === "FLOAT") {
			if (isModeAvailable === newValue) {
				debugLogDetailed("New Value: ", newValue);
				debugLogDetailed("old Value: ", isModeAvailable);
				debugLogDetailed("Value is already available");
				return false;
			} else {
				existingVariable.setValueForMode(mode, newValue);
				debugLogDetailed("Value updated");
				return true;
			}
		} else if (isNotEmpty(isModeAvailable)) {
			if (String(isModeAvailable) == String(newValue)) {
				debugLogDetailed("Value is already available");
				return false;
			} else {
				existingVariable.setValueForMode(mode, newValue);
				debugLogDetailed("Value updated");
				return true;
			}
		} else {
			debugLogDetailed(`Variable found: ${existingVariable.id}`);
			debugLogDetailed("Mode is not available");
			return false;
		}
	} else {
		debugLogDetailed(`Creating variable: ${variableName} in mode: ${mode}`);
		const newVariable = figma.variables.createVariable(
			variableName,
			collection,
			valueType === "STRING" ? "STRING" : "FLOAT"
		);
		newVariable.setValueForMode(mode, newValue);
		debugLogDetailed(`Variable created: ${newVariable.id} in mode: ${mode}`);
		return true;
	}
	return false;
}
function isNotEmpty(value: any): boolean {
	return value !== undefined && value !== null && value !== "";
}

async function getExistingCollectionsAndModes() {
	const collections: {
		[key: string]: {
			name: string;
			id: string;
			defaultModeId: string;
			modes: any[];
		};
	} = (await figma.variables.getLocalVariableCollectionsAsync()).reduce(
		(
			into: { [key: string]: any },
			collection: {
				name: string;
				id: string;
				defaultModeId: string;
				modes: any[];
			}
		) => {
			into[collection.name] = {
				name: collection.name,
				id: collection.id,
				defaultModeId: collection.defaultModeId,
				modes: collection.modes,
			};
			return into;
		},
		{}
	);
	// Log each record in the collections with name and details
	for (const key in collections) {
		if (collections.hasOwnProperty(key)) {
			debugLogDetailed(`Collection Name: ${key}`);
			debugLogDetailed("Details:", collections[key]);
		}
	}
}

function debugLogBasic(...args: any[]) {
	if (DEBUG_BASIC) {
		// console.log(...args);
		const messages = args
			.map((arg) => (typeof arg === "object" ? JSON.stringify(arg) : arg))
			.join("\n");
			sendMessages(messages);
		// figma.ui.postMessage({ type: "debug", message: args.join(" ") });
	}
}

function debugLogDetailed(...args: any[]) {
	if (DEBUG_DETAILED) {
		// console.log(...args);
		const messages = args
			.map((arg) => (typeof arg === "object" ? JSON.stringify(arg) : arg))
			.join("\n");
			sendMessages(messages);
		// figma.ui.postMessage({ type: "debug", message: messages });
	}
}

async function sendMessages(data: string) {
    figma.ui.postMessage({ type: "debug", message: data });
	await delay(100);
}

function delay(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function deleteAllVariableCollections() {
	const collections = figma.variables.getLocalVariableCollections();
	collections.forEach((collection) => {
		collection.remove();
	});
}
