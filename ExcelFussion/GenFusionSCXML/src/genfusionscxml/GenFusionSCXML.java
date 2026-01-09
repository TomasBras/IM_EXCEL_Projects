package genfusionscxml;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import scxmlgen.Fusion.FusionGenerator;
import scxmlgen.Modalities.Gestures;
import scxmlgen.Modalities.Speech;
import scxmlgen.Modalities.Output;

public class GenFusionSCXML {

    private static void ensureConfirmFlowsExcel(String inScxmlFileName, String outScxmlFileName) throws IOException {
        Path inPath = Path.of(inScxmlFileName);
        if (!Files.exists(inPath))
            throw new IOException("SCXML file not found: " + inPath.toAbsolutePath());

        String scxml = Files.readString(inPath, StandardCharsets.UTF_8);

        // Helpers
        class Replace {
            String once(String text, String oldBlock, String newBlock, String label) throws IOException {
                if (!text.contains(oldBlock))
                    throw new IOException("SCXML patch missing block (" + label + ")");
                return text.replace(oldBlock, newBlock);
            }
        }
        Replace replace = new Replace();

        // ------------------------------------------------------------
        // 1) close_excel: require voice confirmation before emitting CLOSE_EXCEL_CONFIRMED
        // ------------------------------------------------------------
        // Speech-only state: immediate timeout to confirmation state.
        scxml = replace.once(
            scxml,
            "  <state id=\"sspeech[speech][close_excel]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[sspeech[speech][close_excel]]\" expr=\"READY\" />\n" +
            "      <assign name=\"data1\" expr=\"${_eventdata.data}\" />\n" +
            "      <send id=\"state1-timer-sspeech[speech][close_excel]\" event=\"timeout-sspeech[speech][close_excel]\" delay=\"5000\" target=\"\" targettype=\"\" namelist=\"\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"main\" event=\"timeout-sspeech[speech][close_excel]\" />\n" +
            "    <transition target=\"sspeech[speech][close_excel]-gestures[gestures][handgrab]\" event=\"[GESTURES][HANDGRAB]\" />\n" +
            "    <onexit>\n" +
            "      <cancel sendid=\"state1-timer-sspeech[speech][close_excel]\" />\n" +
            "    </onexit>\n" +
            "  </state>",
            "  <state id=\"sspeech[speech][close_excel]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[sspeech[speech][close_excel]]\" expr=\"READY\" />\n" +
            "      <assign name=\"data1\" expr=\"${_eventdata.data}\" />\n" +
            "      <send id=\"state1-timer-sspeech[speech][close_excel]\" event=\"timeout-sspeech[speech][close_excel]\" delay=\"0\" target=\"\" targettype=\"\" namelist=\"\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"confirm[close_excel]\" event=\"timeout-sspeech[speech][close_excel]\" />\n" +
            "    <transition target=\"sspeech[speech][close_excel]-gestures[gestures][handgrab]\" event=\"[GESTURES][HANDGRAB]\" />\n" +
            "    <onexit>\n" +
            "      <cancel sendid=\"state1-timer-sspeech[speech][close_excel]\" />\n" +
            "    </onexit>\n" +
            "  </state>",
            "close_excel speech state"
        );

        // Handgrab complementary state: do not emit CLOSE_EXCEL; route into confirmation.
        scxml = replace.once(
            scxml,
            "  <state id=\"sspeech[speech][close_excel]-gestures[gestures][handgrab]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[sspeech[speech][close_excel]-gestures[gestures][handgrab]]\" expr=\"READY\" />\n" +
            "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
            "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[FUSION][CLOSE_EXCEL]')}\" />\n" +
            "      <send event=\"CLOSE_EXCEL\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"main\" />\n" +
            "  </state>",
            "  <state id=\"sspeech[speech][close_excel]-gestures[gestures][handgrab]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[sspeech[speech][close_excel]-gestures[gestures][handgrab]]\" expr=\"READY\" />\n" +
            "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"confirm[close_excel]\" />\n" +
            "  </state>",
            "close_excel handgrab state"
        );

        // Inject the confirmation states for close_excel if missing.
        String closeConfirmStates =
            "  <state id=\"confirm[close_excel]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[confirm[close_excel]]\" expr=\"READY\" />\n" +
            "      <send id=\"state1-timer-confirm[close_excel]\" event=\"timeout-confirm[close_excel]\" delay=\"6000\" target=\"\" targettype=\"\" namelist=\"\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"confirm[close_excel]-confirmed\" event=\"[SPEECH][CONFIRMAR]\" />\n" +
            "    <transition target=\"main\" event=\"[SPEECH][CANCELAR]\" />\n" +
            "    <transition target=\"main\" event=\"timeout-confirm[close_excel]\" />\n" +
            "    <onexit>\n" +
            "      <cancel sendid=\"state1-timer-confirm[close_excel]\" />\n" +
            "    </onexit>\n" +
            "  </state>\n" +
            "  <state id=\"confirm[close_excel]-confirmed\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[confirm[close_excel]-confirmed]\" expr=\"READY\" />\n" +
            "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][CLOSE_EXCEL_CONFIRMED]')}\" />\n" +
            "      <send event=\"CLOSE_EXCEL_CONFIRMED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"main\" />\n" +
            "  </state>\n";

        if (!scxml.contains("<state id=\"confirm[close_excel]\"")) {
            scxml = scxml.replace("</scxml>", closeConfirmStates + "</scxml>");
        }

        // ------------------------------------------------------------
        // 2) undolastaction: require voice confirmation before emitting UNDO_LAST_ACTION_CONFIRMED
        // ------------------------------------------------------------
        scxml = replace.once(
            scxml,
            "  <state id=\"gestures[gestures][undolastaction]-final\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[gestures[gestures][undolastaction]-final]\" expr=\"READY\" />\n" +
            "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][UNDO_LAST_ACTION]')}\" />\n" +
            "      <send event=\"UNDO_LAST_ACTION\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"main\" />\n" +
            "  </state>",
            "  <state id=\"gestures[gestures][undolastaction]-final\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[gestures[gestures][undolastaction]-final]\" expr=\"READY\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"confirm[undo_last_action]\" />\n" +
            "  </state>",
            "undo_last_action final state"
        );

        String undoConfirmStates =
            "  <state id=\"confirm[undo_last_action]\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[confirm[undo_last_action]]\" expr=\"READY\" />\n" +
            "      <send id=\"state1-timer-confirm[undo_last_action]\" event=\"timeout-confirm[undo_last_action]\" delay=\"6000\" target=\"\" targettype=\"\" namelist=\"\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"confirm[undo_last_action]-confirmed\" event=\"[SPEECH][CONFIRMAR]\" />\n" +
            "    <transition target=\"main\" event=\"[SPEECH][CANCELAR]\" />\n" +
            "    <transition target=\"main\" event=\"timeout-confirm[undo_last_action]\" />\n" +
            "    <onexit>\n" +
            "      <cancel sendid=\"state1-timer-confirm[undo_last_action]\" />\n" +
            "    </onexit>\n" +
            "  </state>\n" +
            "  <state id=\"confirm[undo_last_action]-confirmed\">\n" +
            "    <onentry>\n" +
            "      <log label=\"[confirm[undo_last_action]-confirmed]\" expr=\"READY\" />\n" +
            "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][UNDO_LAST_ACTION_CONFIRMED]')}\" />\n" +
            "      <send event=\"UNDO_LAST_ACTION_CONFIRMED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
            "    </onentry>\n" +
            "    <transition target=\"main\" />\n" +
            "  </state>\n";
        if (!scxml.contains("<state id=\"confirm[undo_last_action]\"")) {
            scxml = scxml.replace("</scxml>", undoConfirmStates + "</scxml>");
        }

        // ------------------------------------------------------------
        // 3) Voice-only risky commands: route into a confirmation state
        // ------------------------------------------------------------
        String[] riskySpeech = new String[] { "guardar_ficheiro", "apagar_todos_graficos", "atualizar_notas" };
        for (String cmd : riskySpeech) {
            String timerLine = "      <send id=\"state1-timer-sspeech[speech][" + cmd + "]\" event=\"timeout-sspeech[speech][" + cmd + "]\" delay=\"5000\" target=\"\" targettype=\"\" namelist=\"\" />";
            if (scxml.contains(timerLine)) {
                scxml = scxml.replace(timerLine, timerLine.replace("delay=\"5000\"", "delay=\"0\""));
            }

            String transitionToFinal = "    <transition target=\"speech[speech][" + cmd + "]-final\" event=\"timeout-sspeech[speech][" + cmd + "]\" />";
            if (scxml.contains(transitionToFinal)) {
                scxml = scxml.replace(transitionToFinal, "    <transition target=\"confirm[" + cmd + "]\" event=\"timeout-sspeech[speech][" + cmd + "]\" />");
            } else {
                // If the generator ever changes the shape, fail explicitly.
                throw new IOException("SCXML patch missing transition for sspeech[speech][" + cmd + "]");
            }
        }

        // Inject confirm states for each voice-only risky command.
        class ConfirmState {
            final String cmd;
            final String confirmedEvent;
            ConfirmState(String cmd, String confirmedEvent) { this.cmd = cmd; this.confirmedEvent = confirmedEvent; }
        }

        ConfirmState[] confirmStates = new ConfirmState[] {
            new ConfirmState("guardar_ficheiro", "GUARDAR_FICHEIRO_CONFIRMED"),
            new ConfirmState("apagar_todos_graficos", "APAGAR_TODOS_GRAFICOS_CONFIRMED"),
            new ConfirmState("atualizar_notas", "ATUALIZAR_NOTAS_CONFIRMED"),
        };

        for (ConfirmState cs : confirmStates) {
            if (scxml.contains("<state id=\"confirm[" + cs.cmd + "]\""))
                continue;

            String block =
                "  <state id=\"confirm[" + cs.cmd + "]\">\n" +
                "    <onentry>\n" +
                "      <log label=\"[confirm[" + cs.cmd + "]]\" expr=\"READY\" />\n" +
                "      <send id=\"state1-timer-confirm[" + cs.cmd + "]\" event=\"timeout-confirm[" + cs.cmd + "]\" delay=\"6000\" target=\"\" targettype=\"\" namelist=\"\" />\n" +
                "    </onentry>\n" +
                "    <transition target=\"confirm[" + cs.cmd + "]-confirmed\" event=\"[SPEECH][CONFIRMAR]\" />\n" +
                "    <transition target=\"main\" event=\"[SPEECH][CANCELAR]\" />\n" +
                "    <transition target=\"main\" event=\"timeout-confirm[" + cs.cmd + "]\" />\n" +
                "    <onexit>\n" +
                "      <cancel sendid=\"state1-timer-confirm[" + cs.cmd + "]\" />\n" +
                "    </onexit>\n" +
                "  </state>\n" +
                "  <state id=\"confirm[" + cs.cmd + "]-confirmed\">\n" +
                "    <onentry>\n" +
                "      <log label=\"[confirm[" + cs.cmd + "]-confirmed]\" expr=\"READY\" />\n" +
                "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][" + cs.confirmedEvent + "]')}\" />\n" +
                "      <send event=\"" + cs.confirmedEvent + "\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
                "    </onentry>\n" +
                "    <transition target=\"main\" />\n" +
                "  </state>\n";

            scxml = scxml.replace("</scxml>", block + "</scxml>");
        }

        // Write out
        Files.writeString(Path.of(outScxmlFileName), scxml, StandardCharsets.UTF_8);
    }

    private static void ensureKinectCompatExcel(String scxmlFileName) throws IOException {
    // In runtime logs, Kinect gestures may arrive as numeric id + label (e.g. [3][insertcolumn]).
    // The generator emits symbolic events like [GESTURES][INSERTCOLUMN], so we inject compat transitions.
        Path path = Path.of(scxmlFileName);
        if (!Files.exists(path))
            throw new IOException("SCXML file not found: " + path.toAbsolutePath());

        String scxml = Files.readString(path, StandardCharsets.UTF_8);

    // 1) main: accept numeric gesture ids
    String mainHandgrab = "    <transition target=\"sgestures[gestures][handgrab]\" event=\"[GESTURES][HANDGRAB]\" />";
    if (scxml.contains(mainHandgrab) && !scxml.contains("event=\"[3][insertcolumn]\"")) {
        String numericBlock = mainHandgrab + "\n" +
            "\n    <!-- Compat: Kinect sometimes arrives as numeric id + label (e.g. [3][insertcolumn]) -->" +
            "\n    <transition target=\"sgestures[gestures][calculateaverage]\" event=\"[0][calculateaverage]\" />" +
            "\n    <transition target=\"sgestures[gestures][handgrab]\" event=\"[2][handgrab]\" />" +
            "\n    <transition target=\"sgestures[gestures][insertcolumn]\" event=\"[3][insertcolumn]\" />" +
            "\n    <transition target=\"sgestures[gestures][studentsapproved]\" event=\"[4][studentsapproved]\" />" +
            "\n    <transition target=\"sgestures[gestures][studentsfailed]\" event=\"[5][studentsfailed]\" />" +
            "\n    <transition target=\"sgestures[gestures][swipedown]\" event=\"[6][swipedown]\" />" +
            "\n    <transition target=\"sgestures[gestures][swipeleft]\" event=\"[7][swipeleft]\" />" +
            "\n    <transition target=\"sgestures[gestures][swiperight]\" event=\"[8][swiperight]\" />" +
            "\n    <transition target=\"sgestures[gestures][swipeup]\" event=\"[9][swipeup]\" />" +
            "\n    <transition target=\"sgestures[gestures][undolastaction]\" event=\"[10][undolastaction]\" />" +
            "\n    <transition target=\"sgestures[gestures][zoomin]\" event=\"[11][zoomin]\" />" +
            "\n    <transition target=\"sgestures[gestures][zoomout]\" event=\"[12][zoomout]\" />";
        scxml = scxml.replace(mainHandgrab, numericBlock);
    }

    // 2) INSERIR_COLUNAS: add numeric variants for insertcolumn/approved/failed if the symbolic transitions exist.
    String inserirInsertSymbolic =
        "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][insertcolumn]\" event=\"[GESTURES][INSERTCOLUMN]\" />";
    if (scxml.contains(inserirInsertSymbolic) && !scxml.contains("sspeech[speech][inserir_colunas]-gestures[gestures][insertcolumn]\" event=\"[3][insertcolumn]\"")) {
        scxml = scxml.replace(inserirInsertSymbolic,
            inserirInsertSymbolic + "\n" +
            "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][insertcolumn]\" event=\"[3][insertcolumn]\" />");
    }

    String inserirApprovedSymbolic =
        "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][studentsapproved]\" event=\"[GESTURES][STUDENTSAPPROVED]\" />";
    if (scxml.contains(inserirApprovedSymbolic) && !scxml.contains("sspeech[speech][inserir_colunas]-gestures[gestures][studentsapproved]\" event=\"[4][studentsapproved]\"")) {
        scxml = scxml.replace(inserirApprovedSymbolic,
            inserirApprovedSymbolic + "\n" +
            "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][studentsapproved]\" event=\"[4][studentsapproved]\" />");
    }

    String inserirFailedSymbolic =
        "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][studentsfailed]\" event=\"[GESTURES][STUDENTSFAILED]\" />";
    if (scxml.contains(inserirFailedSymbolic) && !scxml.contains("sspeech[speech][inserir_colunas]-gestures[gestures][studentsfailed]\" event=\"[5][studentsfailed]\"")) {
        scxml = scxml.replace(inserirFailedSymbolic,
            inserirFailedSymbolic + "\n" +
            "    <transition target=\"sspeech[speech][inserir_colunas]-gestures[gestures][studentsfailed]\" event=\"[5][studentsfailed]\" />");
    }

    // 3) DESTACAR_APROVADOS_REPROVADOS: add numeric variants for approved/failed.
    String destacarApprovedSymbolic =
        "    <transition target=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]\" event=\"[GESTURES][STUDENTSAPPROVED]\" />";
    if (scxml.contains(destacarApprovedSymbolic) && !scxml.contains("sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]\" event=\"[4][studentsapproved]\"")) {
        scxml = scxml.replace(destacarApprovedSymbolic,
            destacarApprovedSymbolic + "\n" +
            "    <transition target=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]\" event=\"[4][studentsapproved]\" />");
    }

    String destacarFailedSymbolic =
        "    <transition target=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]\" event=\"[GESTURES][STUDENTSFAILED]\" />";
    if (scxml.contains(destacarFailedSymbolic) && !scxml.contains("sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]\" event=\"[5][studentsfailed]\"")) {
        scxml = scxml.replace(destacarFailedSymbolic,
            destacarFailedSymbolic + "\n" +
            "    <transition target=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]\" event=\"[5][studentsfailed]\" />");
    }

    // 3b) Ensure gesture-only approved/failed outputs are emitted as gesture events (not HIGHLIGHT_RESULTS).
    // Some generator configurations historically emitted HIGHLIGHT_RESULTS for these gestures; enforce the intended mapping here.
    scxml = scxml.replace(
        "  <state id=\"gestures[gestures][studentsapproved]-final\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[gestures[gestures][studentsapproved]-final]\" expr=\"READY\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"gestures[gestures][studentsapproved]-final\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[gestures[gestures][studentsapproved]-final]\" expr=\"READY\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[GESTURES][STUDENTSAPPROVED]')}\" />\n" +
        "      <send event=\"STUDENTSAPPROVED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

    scxml = scxml.replace(
        "  <state id=\"gestures[gestures][studentsfailed]-final\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[gestures[gestures][studentsfailed]-final]\" expr=\"READY\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"gestures[gestures][studentsfailed]-final\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[gestures[gestures][studentsfailed]-final]\" expr=\"READY\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF2(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1, '[GESTURES][STUDENTSFAILED]')}\" />\n" +
        "      <send event=\"STUDENTSFAILED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

    // When voice+gesture happens together for 'destacar...', prefer the gesture specificity.
    scxml = scxml.replace(
        "  <state id=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsapproved]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[GESTURES][STUDENTSAPPROVED]')}\" />\n" +
        "      <send event=\"STUDENTSAPPROVED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

    scxml = scxml.replace(
        "  <state id=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sspeech[speech][destacar_aprovados_reprovados]-gestures[gestures][studentsfailed]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[GESTURES][STUDENTSFAILED]')}\" />\n" +
        "      <send event=\"STUDENTSFAILED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

    scxml = scxml.replace(
        "  <state id=\"sgestures[gestures][studentsapproved]-speech[speech][destacar_aprovados_reprovados]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sgestures[gestures][studentsapproved]-speech[speech][destacar_aprovados_reprovados]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"sgestures[gestures][studentsapproved]-speech[speech][destacar_aprovados_reprovados]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sgestures[gestures][studentsapproved]-speech[speech][destacar_aprovados_reprovados]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[GESTURES][STUDENTSAPPROVED]')}\" />\n" +
        "      <send event=\"STUDENTSAPPROVED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

    scxml = scxml.replace(
        "  <state id=\"sgestures[gestures][studentsfailed]-speech[speech][destacar_aprovados_reprovados]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sgestures[gestures][studentsfailed]-speech[speech][destacar_aprovados_reprovados]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[FUSION][HIGHLIGHT_RESULTS]')}\" />\n" +
        "      <send event=\"HIGHLIGHT_RESULTS\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>",
        "  <state id=\"sgestures[gestures][studentsfailed]-speech[speech][destacar_aprovados_reprovados]\">\n" +
        "    <onentry>\n" +
        "      <log label=\"[sgestures[gestures][studentsfailed]-speech[speech][destacar_aprovados_reprovados]]\" expr=\"READY\" />\n" +
        "      <assign name=\"data2\" expr=\"${_eventdata.data}\" />\n" +
        "      <commons:var name=\"newExtensionNotification\" expr=\"${mmi:newExtensionNotificationF(contextId, 'FUSION', 'IM', mmi:newRequestId(contextId), null, data1,data2, '[GESTURES][STUDENTSFAILED]')}\" />\n" +
        "      <send event=\"STUDENTSFAILED\" target=\"IM\" targettype=\"MC\" namelist=\"newExtensionNotification\" />\n" +
        "    </onentry>"
    );

   
        String[] speechStates = new String[] {
                "calcular_media",
        "inserir_colunas",
                "criar_pivot_table",
                "gerar_grafico_turma",
                "gerar_grafico_barras_aluno",
                "gerar_grafico_perguntas_t2",
                "close_excel"
        };

        for (String speechState : speechStates) {
            String symbolicLine =
                    "    <transition target=\"sspeech[speech][" + speechState + "]-gestures[gestures][handgrab]\" event=\"[GESTURES][HANDGRAB]\" />";
            String compatLine =
                    symbolicLine + "\n" +
                    "    <!-- Compat: Kinect numeric id for handgrab -->\n" +
                    "    <transition target=\"sspeech[speech][" + speechState + "]-gestures[gestures][handgrab]\" event=\"[2][handgrab]\" />";

            if (scxml.contains(symbolicLine) && !scxml.contains("sspeech[speech][" + speechState + "]-gestures[gestures][handgrab]\" event=\"[2][handgrab]")) {
                scxml = scxml.replace(symbolicLine, compatLine);
            }
        }

        Files.writeString(path, scxml, StandardCharsets.UTF_8);
    }

    public static void main(String[] args) throws IOException {

        FusionGenerator fg = new FusionGenerator();

    
        fg.Redundancy(Speech.CALCULAR_MEDIA, Gestures.CALCULATEAVERAGE, Output.CALCULATE_AVERAGE);
        fg.Redundancy(Speech.INSERIR_COLUNAS, Gestures.INSERTCOLUMN, Output.INSERT_COLUMN);

  
        fg.Single(Speech.CRIAR_PIVOT_TABLE, Output.CREATE_PIVOT);
        fg.Single(Speech.GERAR_GRAFICO_TURMA, Output.GENERATE_GRAPH_TURMA);
        fg.Single(Speech.GERAR_GRAFICO_BARRAS_ALUNO, Output.GENERATE_GRAPH_ALUNO);
        fg.Single(Speech.GERAR_GRAFICO_PERGUNTAS_T2, Output.GENERATE_GRAPH_PERGUNTAS_T2);

        fg.Single(Speech.OPERACOES_MATEMATICAS, Output.OPERACOES_MATEMATICAS);
        fg.Single(Speech.APAGAR_TODOS_GRAFICOS, Output.APAGAR_TODOS_GRAFICOS);
        fg.Single(Speech.ATUALIZAR_NOTAS, Output.ATUALIZAR_NOTAS);
        fg.Single(Speech.GUARDAR_FICHEIRO, Output.GUARDAR_FICHEIRO);
        fg.Single(Speech.HELPER, Output.HELPER);

        // Encaminha comandos de confirmação por voz (útil para complementaridade/diálogo).
        fg.Single(Speech.CONFIRMAR, Output.CONFIRMAR);
        fg.Single(Speech.CANCELAR, Output.CANCELAR);

    
        // Complementares (seleção via handgrab)
        fg.Complementary(Speech.CALCULAR_MEDIA, Gestures.HANDGRAB, Output.CALCULATE_AVERAGE_ON_SELECTION);
        fg.Complementary(Speech.GERAR_GRAFICO_TURMA, Gestures.HANDGRAB, Output.GENERATE_GRAPH_TURMA_ON_SELECTION);
        fg.Complementary(Speech.GERAR_GRAFICO_BARRAS_ALUNO, Gestures.HANDGRAB, Output.GENERATE_GRAPH_ALUNO_ON_SELECTION);

        // Complementar (novo): gesto insertcolumn + voz melhoria_possivel

        fg.Complementary(Speech.INSERIR_COLUNAS, Gestures.STUDENTSAPPROVED, Output.INSERT_COLUMN_THEN_HIGHLIGHT_APPROVED);
        fg.Complementary(Speech.INSERIR_COLUNAS, Gestures.STUDENTSFAILED, Output.INSERT_COLUMN_THEN_HIGHLIGHT_FAILED);


    // Permite fechar o Excel só com voz (sem exigir handgrab).
    fg.Single(Speech.CLOSE_EXCEL, Output.CLOSE_EXCEL);
    // Mantém também a variante complementar com handgrab.
    fg.Complementary(Speech.CLOSE_EXCEL, Gestures.HANDGRAB, Output.CLOSE_EXCEL);

      
        fg.Single(Gestures.SWIPELEFT, Output.SWIPE_LEFT);
        fg.Single(Gestures.SWIPERIGHT, Output.SWIPE_RIGHT);
        fg.Single(Gestures.SWIPEUP, Output.SWIPE_UP);
        fg.Single(Gestures.SWIPEDOWN, Output.SWIPE_DOWN);

       
        fg.Single(Gestures.STUDENTSAPPROVED, Output.STUDENTSAPPROVED);
        fg.Single(Gestures.STUDENTSFAILED, Output.STUDENTSFAILED);

       
        fg.Single(Gestures.ZOOMIN, Output.ZOOM_IN);
        fg.Single(Gestures.ZOOMOUT, Output.ZOOM_OUT);

     
        fg.Single(Gestures.UNDOLASTACTION, Output.UNDO_LAST_ACTION);

        
        // Generate base SCXML, then post-process it so confirmations are preserved across regenerations.
        fg.Build("fusion_excel_raw.scxml");
        ensureConfirmFlowsExcel("fusion_excel_raw.scxml", "fusion_excel.scxml");
        ensureKinectCompatExcel("fusion_excel.scxml");
        System.out.println("Ficheiro SCXML do Excel gerado com sucesso!");
    }
}
