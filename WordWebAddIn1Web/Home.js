
(function () {
    "use strict";
    //Test for fork Syncing          
    var messageBanner;
    
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                return;
            }
            document.getElementById("JDtextarea").addEventListener("input", function () {
                segregateData(document.getElementById("JDtextarea").value);
            });
             //Add a click event handler for the highlight button.

            document.getElementById("AnalyzeJD").addEventListener("click", function () {
                handle_stackAnalysis(frontEndStacks, found_FrontEndStacks);
                handle_stackAnalysis(secondaryBackEndStacks, found_secondaryBackEndStacks);
                handle_stackAnalysis(styleFrameworks, found_styleFrameworks);
                handle_stackAnalysis(sqlDatabases, found_sqlDatabases);
                handle_stackAnalysis(noSDatabases, found_noSDatabases);
                handle_stackAnalysis(planningTools, found_planningTools);
                handle_stackAnalysis(developmentTools, found_developmentTools);
                handle_stackAnalysis(versionControl, found_versionControl);
                handle_stackAnalysis(buildTools, found_buildTools);
                handle_stackAnalysis(testingTools, found_testingTools);
                handle_stackAnalysis(deployementTools, found_deployementTools);
                handle_stackAnalysis(operationsTools, found_operationsTools);
                handle_stackAnalysis(cloudServices, found_cloudServices);
                handle_stackAnalysis(mlSkills, found_MLSkills);
                handle_stackAnalysis(dataOnly, found_dataOnly);
                handle_stackAnalysis(mlDataToolsFrameworks, found_mlDataToolsFrameworks);
                handle_stackAnalysis(bigData, found_bigData);
                handle_stackAnalysis(blockChain, found_blockChain);
                handle_stackAnalysis(others, found_others);
            });

            Array.from(document.getElementsByClassName("replaceKeywordsButton")).forEach(value => {
                value.addEventListener("click", function () {
                    replaceTextWithRegex(string_frontEndStacks, Array.from(document.getElementById("frontEndStacks_checkboxContainer").querySelectorAll("input[type=checkbox][name='frontEndStacks_checkboxContainer-checkbox-group[]']:checked"), e => e.value), frontEndStacks);
                    replaceTextWithRegex(string_primaryBackEndStacks, Array.from(document.getElementById("primaryBackEndStacks_checkboxContainer").querySelectorAll("input[type=checkbox][name='primaryBackEndStacks_checkboxContainer-checkbox-group[]']:checked"), e => e.value), primaryBackEndStacks);
                    replaceTextWithRegex(string_secondaryBackEndStacks, Array.from(document.getElementById("secondaryBackEndStacks_checkboxContainer").querySelectorAll("input[type=checkbox][name='secondaryBackEndStacks_checkboxContainer-checkbox-group[]']:checked"), e => e.value), secondaryBackEndStacks);
                    replaceTextWithRegex(string_styleFrameworks, Array.from(document.getElementById("styleFrameworks_checkboxContainer").querySelectorAll("input[type=checkbox][name='styleFrameworks_checkboxContainer-checkbox-group[]']:checked"), e => e.value), styleFrameworks);
                    replaceTextWithRegex(string_sqlDatabases, Array.from(document.getElementById("sqlDatabases_checkboxContainer").querySelectorAll("input[type=checkbox][name='sqlDatabases_checkboxContainer-checkbox-group[]']:checked"), e => e.value), sqlDatabases);
                    replaceTextWithRegex(string_noSDatabases, Array.from(document.getElementById("noSDatabases_checkboxContainer").querySelectorAll("input[type=checkbox][name='noSDatabases_checkboxContainer-checkbox-group[]']:checked"), e => e.value), noSDatabases);
                    replaceTextWithRegex(string_planningTools, Array.from(document.getElementById("planningTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='planningTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), planningTools);
                    replaceTextWithRegex(string_developmentTools, Array.from(document.getElementById("developmentTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='developmentTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), developmentTools);
                    replaceTextWithRegex(string_versionControl, Array.from(document.getElementById("versionControl_checkboxContainer").querySelectorAll("input[type=checkbox][name='versionControl_checkboxContainer-checkbox-group[]']:checked"), e => e.value), versionControl);
                    replaceTextWithRegex(string_buildTools, Array.from(document.getElementById("buildTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='buildTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), buildTools);
                    replaceTextWithRegex(string_testingTools, Array.from(document.getElementById("testingTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='testingTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), testingTools);
                    replaceTextWithRegex(string_deployementTools, Array.from(document.getElementById("deployementTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='deployementTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), deployementTools);
                    replaceTextWithRegex(string_operationsTools, Array.from(document.getElementById("operationsTools_checkboxContainer").querySelectorAll("input[type=checkbox][name='operationsTools_checkboxContainer-checkbox-group[]']:checked"), e => e.value), operationsTools);
                    replaceTextWithRegex(string_cloudServices, Array.from(document.getElementById("cloudServices_checkboxContainer").querySelectorAll("input[type=checkbox][name='cloudServices_checkboxContainer-checkbox-group[]']:checked"), e => e.value), cloudServices);
                    replaceTextWithRegex(string_mlSkills, Array.from(document.getElementById("mlSkills_checkboxContainer").querySelectorAll("input[type=checkbox][name='mlSkills_checkboxContainer-checkbox-group[]']:checked"), e => e.value), mlSkills);
                    replaceTextWithRegex(string_dataOnly, Array.from(document.getElementById("dataOnly_checkboxContainer").querySelectorAll("input[type=checkbox][name='dataOnly_checkboxContainer-checkbox-group[]']:checked"), e => e.value), dataOnly);
                    replaceTextWithRegex(string_mlDataToolsFrameworks, Array.from(document.getElementById("mlDataToolsFrameworks_checkboxContainer").querySelectorAll("input[type=checkbox][name='mlDataToolsFrameworks_checkboxContainer-checkbox-group[]']:checked"), e => e.value), mlDataToolsFrameworks);
                    replaceTextWithRegex(string_bigData, Array.from(document.getElementById("bigData_checkboxContainer").querySelectorAll("input[type=checkbox][name='bigData_checkboxContainer-checkbox-group[]']:checked"), e => e.value), bigData);
                    replaceTextWithRegex(string_blockChain, Array.from(document.getElementById("blockChain_checkboxContainer").querySelectorAll("input[type=checkbox][name='blockChain_checkboxContainer-checkbox-group[]']:checked"), e => e.value), blockChain);
                    replaceTextWithRegex(string_others, Array.from(document.getElementById("others_checkboxContainer").querySelectorAll("input[type=checkbox][name='others_checkboxContainer-checkbox-group[]']:checked"), e => e.value), others);
                    copy();
                });
            });

            Array.from(document.getElementsByClassName("clearSelection")).forEach(value => {
                value.addEventListener("click", function () {
                    //clear all checkboxes                    
                    var checkboxes = document.querySelectorAll("input[type='checkbox']");
                    // Loop through the checkboxes and uncheck them
                    checkboxes.forEach(function (checkbox) {
                        checkbox.checked = false;
                    });
                });
            });

            Array.from(document.getElementsByClassName("copyName")).forEach(value => {
                value.addEventListener("click", function () {
                    copy();
                });
            });


        });
    };
    function segregateData(data) {
        company = data.split("^")[0];
        jobTitle = data.split("^")[1];
        jobLink = data.split("^")[2];
        jobDescription = data.split("^")[3];
        document.getElementById("company").value = company;
        document.getElementById("jobTitle").value = jobTitle;
        document.getElementById("jobLink").value = jobLink;
        document.getElementById("jobDescription").value = jobDescription;
    }
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();

let rawJDString = '';
let jobTitle = '';
let jobLink = '';
let company = '';
let jobDescription = '';

const string_progLanguages = "progLanguages";
//const regex_progLanguages = "/progLanguages/g";
let found_ProgLanguages = [];
const progLanguages = [, , , ];

const string_frontEndStacks = "frontEndStacks";
//const regex_frontEndStacks = "/frontEndStacks/g";
let found_FrontEndStacks = [];
const frontEndStacks = ["React", "Vue", "Angular", "Ember", "GraphQL", "Apollo", "Redux", "jQuery", "TypeScript", "JavaScript","HTML", "CSS"];

const string_primaryBackEndStacks = "primaryBackEndStacks";
//const regex_primaryBackEndStacks = "/primaryBackEndStacks/g";
let found_primaryBackEndStacks = [];
const primaryBackEndStacks = ["Ruby on Rails", "Python", "PHP", "Laravel", "NestJs", "NodeJS", "ExpressJS",];

const string_secondaryBackEndStacks = "secondaryBackEndStacks";
//const regex_secondaryBackEndStacks = "/secondaryBackEndStacks/g";
let found_secondaryBackEndStacks = [];
const secondaryBackEndStacks = ["Rails", "Django", "Flask", "Laravel" ,"Node"];

const string_styleFrameworks = "styleFrameworks";
//const regex_styleFrameworks = "/styleFrameworks/g";
let found_styleFrameworks = [];
const styleFrameworks = ["Bootstrap", "Tailwind", "MaterialUI"];

const string_sqlDatabases = "sqlDatabases";
//const regex_sqlDatabases = "/sqlDatabases/g";
let found_sqlDatabases = [];
const sqlDatabases = ["SQLite", "Oracle", "Microsoft SQL Server", "MySQL", "PostgreSQL"];

const string_noSDatabases = "noSDatabases";
//const regex_noSDatabases = "/noSDatabases/g";
let found_noSDatabases = [];
const noSDatabases = ["MongoDB", "DynamoDB", "Redis", "Cassandra"];

const string_planningTools = "planningTools";
//const regex_planningTools = "/planningTools/g";
let found_planningTools = [];
const planningTools = ["Clickup", "Asana", "JIRA", "Trello", "RedShift", "PivotalTracker", "BaseCamp"];

const string_developmentTools = "developmentTools";
//const regex_developmentTools = "/developmentTools/g";
let found_developmentTools = [];
const developmentTools = ["Visual Studio", "VS Code", "Postman"];

const string_versionControl = "versionControl";
//const regex_versionControl = "/versionControl/g";
let found_versionControl = [];
const versionControl = ["Github", "Gitlab", "BitBucket", "Git", "GitKraken"];

const string_buildTools = "buildTools";
//const regex_buildTools = "/buildTools/g";
let found_buildTools = [];
const buildTools = ["Jenkins", "Maven", "Gradle", "Github Actions", "CircleCI"];

//const regex_testingTools = "/testingTools/g";
const string_testingTools = "testingTools";
let found_testingTools = [];
const testingTools = ["Selenium", "JUnit", "Jest", "Chai", "Mocha", "Rspec", "RTL"];

const string_deployementTools = "deployementTools";
//const regex_deployementTools = "/deployementTools/g";
let found_deployementTools = [];
const deployementTools = ["Docker", "Ansible", "Terraform", "Kubernetes"];

const string_operationsTools = "operationsTools";
//const regex_operationsTools = "/operationsTools/g";
let found_operationsTools = [];
const operationsTools = ["Prometheus", "ELK", "Nagios"];

const string_cloudServices = "cloudServices";
//const regex_cloudServi//ces = "/cloudServices/g";
let found_cloudServices = [];
const cloudServices = ["Netlify", "Heroku", "Firebase", "Azure Cloud", "GCP", "Google Cloud Platform", "AWS", "Amazon Web Services"];

const string_mlSkills = "mlSkills";
//const regex_mlSkills = "/mlSkills/g";
let found_MLSkills = [];
const mlSkills = ["Supervised Learning", "Unsupervised Learning", "Reinforcement Learning", "Decision Trees", "Random Forests", "Support Vector Machines", "SVM", "Natural Language Processing", "NLP", "Computer Vision"];

const string_dataOnly = "dataOnly";
let found_dataOnly = [];
const dataOnly = ["Data Preprocessing", "Data Visualization", "Feature Engineering", "Data Analysis", "Data Cleaning"];

const string_mlDataToolsFrameworks = "mlDataToolsFrameworks";
//const regex_mlDataToolsFrameworks = "/mlDataToolsFrameworks/g";
let found_mlDataToolsFrameworks = [];
const mlDataToolsFrameworks = ["TensorFlow", "PyTorch", "Scikit", "Keras", "Pandas", "NumPy", "OpenCV", "NLTK", "spaCy"];

const string_bigData = "bigData";
//const regex_bigDataSkills = "/bigData/g";
let found_bigData = [];
const bigData = ["Hadoop", "Spark", "MapReduce", "Apache Airflow", "Apache Kafka", "AWS Glue", "EMR", "RedShift"];

const string_blockChain = "blockChain";
//const regex_blockChain = "/blockChain/g";
let found_blockChain = [];
const blockChain = ["Solidity", "Smart Contracts", "DeFi", "Dapps", "Exchanges", "Token Development", "Security Audits", "NFT", "wallets", "Aggregators", "Signature verifications"];

const string_others = "othestring";
//const regex_others = "/othersReg/g";
let found_others = [];
const others = ["Shopify", "Wordpress"];

function makeCheckBoxes(itemsList, containerId, label) {
    var checkboxGroup = document.createElement("div");
    checkboxGroup.className = "formbuilder-checkbox-group form-group field-checkbox-group-1696514414261 shadow p-3 bg-white rounded";
    checkboxGroup.style = "width: 270px;"
    checkboxGroup.id = containerId + "_group";

    var labelElement = document.createElement("label");
    labelElement.setAttribute("for", `checkbox-group-${containerId}`);
    labelElement.className = "formbuilder-checkbox-group-label h7 text-primary";
    labelElement.textContent = label;

    var checkboxContainer = document.createElement("div");
    checkboxContainer.className = "checkbox-group";
    checkboxContainer.id = containerId;

    itemsList.forEach(function (tech, index) {
        var checkbox = document.createElement("div");
        checkbox.className = "formbuilder-checkbox";
        checkbox.innerHTML = `
        <input name="${containerId}-checkbox-group[]" access="false" id="checkbox-group-${index}" value="${tech}" type="checkbox">
        <label for="checkbox-group-${index}">${tech}</label>`;
        checkboxContainer.appendChild(checkbox);
    });

    checkboxGroup.appendChild(labelElement);
    checkboxGroup.appendChild(checkboxContainer);

    var parentContainer = document.getElementById("checkboxesParent");
    parentContainer.appendChild(checkboxGroup);
}

function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function handle_stackAnalysis(modelArray, foundArray) {
    const regexTerms = modelArray.map((term) => `\\b${escapeRegExp(term.trim())}\\b`).join("|");
    const regex = new RegExp(regexTerms, "gi"); // add the i flag here
    const matches = jobDescription.match(regex);
    if (matches) {
        matches.forEach((match) => {
            foundArray.push(modelArray.find((element) => match.includes(element)));
        });
    }
    if (foundArray.length===0) {
        foundArray.push(modelArray[Math.floor(Math.random() * modelArray.length)]);
    }
    checkCheckboxes(foundArray);
}
function checkCheckboxes(foundArr) {
    foundArr.forEach(function (match) {
        console.log(foundArr);
        var checkbox = document.querySelector(`[value="${match}"]`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
}
function replaceTextWithRegex(replacementString, foundArray, modelArray) {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {
        var body = context.document.body;
        // Queue a command to search the document body for the word "hello".
        var searchResults = body.search(replacementString, { matchCase: false, matchWholeWord: false });
        // Queue a command to load the results.
        context.load(searchResults, "text");
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {

            // Loop through the results and highlight each one with yellow color.
            for (var i = 0; i < searchResults.items.length; i++) {
                searchResults.items[i].clear();
                searchResults.items[i].insertText(foundArray.length > 0 ? foundArray[Math.floor(Math.random() * foundArray.length)] : modelArray[Math.floor(Math.random() * modelArray.length)]);
                searchResults.items[i].font.highlightColor = '#FFFF00';
            }
            // Synchronize again to apply the changes.
            return context.sync();
        });
    });
    //    .catch(function (error) {
    //    throw
    //    // Handle any errors that occurred.
    //    //console.log("Error: " + JSON.stringify(error));
    //    //if (error instanceof OfficeExtension.Error) {
    //    //    console.log("Debug info: " + JSON.stringify(error.debugInfo));
    //    //}
    //});

}
function copy() {
    let fileName = "Murtza_" + document.getElementById("company").value + "_" + document.getElementById("jobTitle").value;

    let temp = document.createElement("textarea");
    document.body.appendChild(temp);
    temp.value = fileName;
    temp.select();
    navigator.clipboard.writeText(temp.value);
    temp.remove();
}