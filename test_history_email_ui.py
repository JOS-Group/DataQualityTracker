import unittest
from pathlib import Path


APP_SOURCE = Path(__file__).with_name("app6.py").read_text(encoding="utf-8")


def source_between(start, end):
    return APP_SOURCE.split(start, 1)[1].split(end, 1)[0]


class HistoryEmailUiTests(unittest.TestCase):
    def test_agent_summary_is_rendered_once_outside_snapshot_details(self):
        page = source_between("function historyPage(info,data){", "\nasync function renderByo")
        self.assertIn("const agentPanel=dataset&&data?historyEmailPanel(data):'';", page)
        self.assertIn("${picker}${agentPanel}${content}", page)
        self.assertEqual(page.count("historyEmailPanel(data)"), 1)
        self.assertNotIn("data.comparison?'':`<br>${historyEmailPanel(data)}`", page)

    def test_summary_and_modal_keep_existing_agent_actions_and_delivery_tools(self):
        panel = source_between("function historyEmailPanel(data){", "\nfunction historyDrilldown")
        for expected in (
            'class="panel historyAgentPanel"',
            'id="historyEmailDialog"',
            'aria-labelledby="historyEmailDialogTitle"',
            'aria-label="Saved history email agents"',
            'onclick="newHistoryEmailTemplate()"',
            'onclick="applyHistoryEmailPreset(\'starts_with_1\')"',
            'id="historyEmailPreviewHolder"',
            "Execution History &amp; Delivery Log",
            "editHistoryEmailById(",
            "runHistoryEmailTemplate(",
            "testHistoryEmailTemplate(",
            "deleteHistoryEmailTemplate(",
            "Open file Email Alerts",
        ):
            with self.subTest(expected=expected):
                self.assertIn(expected, panel)

    def test_agent_summary_is_collapsed_by_default_and_remembers_toggle_state(self):
        panel = source_between("function historyEmailPanel(data){", "\nfunction historyDrilldown")
        disclosure = panel.split('<details class="panel historyAgentPanel"', 1)[1].split('</details>', 1)[0]
        self.assertIn("ontoggle=\"state.historyEmailPanelOpen=this.open\"", panel)
        self.assertIn("${state.historyEmailPanelOpen?' open':''}", panel)
        self.assertNotIn(" open>", panel.split("ontoggle=", 1)[0])
        self.assertLess(disclosure.index("<summary class=\"historyAgentSummary\">"), disclosure.index("historyAgentContent"))
        self.assertIn('onclick="newHistoryEmailTemplate()"', disclosure)
        self.assertIn('aria-label="Saved history email agents"', disclosure)
        self.assertLess(panel.index("</details>"), panel.index('<dialog id="historyEmailDialog"'))

    def test_opening_editor_renders_selected_template_before_showing_dialog(self):
        opener = source_between("async function openHistoryEmailEditor(", "\nasync function editHistoryEmailById")
        self.assertLess(opener.index("state.historyEmailEditingId=id||null"), opener.index("await render()"))
        self.assertLess(opener.index("await render()"), opener.index("dialog.showModal()"))
        self.assertIn("document.getElementById('historyEmailName')?.focus", opener)

        editor_actions = source_between("async function editHistoryEmailById(", "\nasync function deleteHistoryEmailTemplate")
        self.assertIn("await openHistoryEmailEditor(id)", editor_actions)
        self.assertIn("await openHistoryEmailEditor()", editor_actions)

    def test_save_rerenders_only_after_success_and_preserves_form_on_error(self):
        save = source_between("async function saveHistoryEmailTemplate(){", "\nasync function previewHistoryEmailTemplate")
        success = save.split("  }catch(err){", 1)[0]
        failure = save.split("  }catch(err){", 1)[1]
        self.assertLess(success.index("await postJson('/api/byo/history/email-agent/template'"), success.index("await render()"))
        self.assertIn("History Email Agent saved successfully", success)
        self.assertNotIn("render()", failure)
        self.assertIn("Could not save History Agent", failure)

    def test_changing_dataset_clears_stale_agent_edit_state_and_file_alert_page_remains(self):
        selection = source_between("async function selectHistoryDataset(value){", "\nfunction historyPeriod")
        self.assertIn("if(value!==state.historyDataset)", selection)
        self.assertIn("state.historyEmailEditingId=null", selection)
        self.assertIn("showPage('byo-email')", APP_SOURCE)
        self.assertIn("if(state.page==='byo-email')emailInfo=await api(`", APP_SOURCE)
        self.assertIn("byoEmailAgentPage(emailInfo||{},info)", APP_SOURCE)


if __name__ == "__main__":
    unittest.main()