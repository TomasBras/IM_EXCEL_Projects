package scxmlgen.Modalities;

import scxmlgen.interfaces.IOutput;

public enum Output implements IOutput {


    // =========================================================
    // CORE (ações base) — existe no FusionEngine runtime
    // =========================================================
    CALCULATE_AVERAGE("[FUSION][CALCULATE_AVERAGE]"),
    INSERT_COLUMN("[FUSION][INSERT_COLUMN]"),
    CREATE_PIVOT("[FUSION][CREATE_PIVOT]"),
    GENERATE_GRAPH_TURMA("[FUSION][GENERATE_GRAPH_TURMA]"),
    GENERATE_GRAPH_ALUNO("[FUSION][GENERATE_GRAPH_ALUNO]"),
    GENERATE_GRAPH_PERGUNTAS_T2("[FUSION][GENERATE_GRAPH_PERGUNTAS_T2]"),
    HIGHLIGHT_RESULTS("[FUSION][HIGHLIGHT_RESULTS]"),

    // =========================================================
    // REDUNDÂNCIA (Speech/Gesture fazem o mesmo) — usado no gerador
    // Ex.: dizer "calcular média" OU fazer gesto calculateaverage
    // =========================================================
    // (output é o mesmo, a redundância é a regra de fusão)

    // =========================================================
    // COMPLEMENTARES (combinações) — existe no FusionEngine runtime
    // - *_ON_SELECTION: handgrab + voz (ou gesto+voz) => operar só na seleção
    // - INSERT_COLUMN_THEN_*: sequência com gesto complementar
    // =========================================================
    CALCULATE_AVERAGE_ON_SELECTION("[FUSION][CALCULATE_AVERAGE_ON_SELECTION]"),
    HIGHLIGHT_RESULTS_ON_SELECTION("[FUSION][HIGHLIGHT_RESULTS_ON_SELECTION]"),
    GENERATE_GRAPH_TURMA_ON_SELECTION("[FUSION][GENERATE_GRAPH_TURMA_ON_SELECTION]"),
    GENERATE_GRAPH_ALUNO_ON_SELECTION("[FUSION][GENERATE_GRAPH_ALUNO_ON_SELECTION]"),

    INSERT_COLUMN_THEN_HIGHLIGHT_APPROVED("[FUSION][INSERT_COLUMN_THEN_HIGHLIGHT_APPROVED]"),
    INSERT_COLUMN_THEN_HIGHLIGHT_FAILED("[FUSION][INSERT_COLUMN_THEN_HIGHLIGHT_FAILED]"),

    // =========================================================
    // CONFIRMAÇÕES (2 passos) — existe no FusionEngine runtime
    // CLOSE_EXCEL / UNDO disparam primeiro o pedido e só depois *_CONFIRMED
    // =========================================================
    CLOSE_EXCEL("[FUSION][CLOSE_EXCEL]"),
    CLOSE_EXCEL_CONFIRMED("[FUSION][CLOSE_EXCEL_CONFIRMED]"),
    UNDO_LAST_ACTION("[FUSION][UNDO_LAST_ACTION]"),
    UNDO_LAST_ACTION_CONFIRMED("[FUSION][UNDO_LAST_ACTION_CONFIRMED]"),

    // =========================================================
    // LEGADO / NÃO USADO NO FusionEngine runtime (hoje)
    // Mantido no gerador porque o IM pode tratar estes intents diretamente.
    // =========================================================
    MELHORIA_REAL("[FUSION][MELHORIA_REAL]"),
    MELHORIA_POSSIVEL("[FUSION][MELHORIA_POSSIVEL]"),
    OPERACOES_MATEMATICAS("[FUSION][OPERACOES_MATEMATICAS]"),
    APAGAR_TODOS_GRAFICOS("[FUSION][APAGAR_TODOS_GRAFICOS]"),
    ATUALIZAR_NOTAS("[FUSION][ATUALIZAR_NOTAS]"),
    GUARDAR_FICHEIRO("[FUSION][GUARDAR_FICHEIRO]"),
    HELPER("[FUSION][HELPER]"),

    // Confirmar/Cancelar como outputs (o FusionEngine runtime apenas consome estes como inputs)
    CONFIRMAR("[FUSION][CONFIRMAR]"),
    CANCELAR("[FUSION][CANCELAR]"),

    // =========================================================
    // PASS-THROUGH (gestos encaminhados para o IM) — existe no FusionEngine runtime
    // =========================================================
    SWIPE_LEFT("[GESTURES][SWIPELEFT]"),
    SWIPE_RIGHT("[GESTURES][SWIPERIGHT]"),
    SWIPE_UP("[GESTURES][SWIPEUP]"),
    SWIPE_DOWN("[GESTURES][SWIPEDOWN]"),

    STUDENTSAPPROVED("[GESTURES][STUDENTSAPPROVED]"),
    STUDENTSFAILED("[GESTURES][STUDENTSFAILED]"),

    ZOOM_IN("[GESTURES][ZOOMIN]"),
    ZOOM_OUT("[GESTURES][ZOOMOUT]");

    private final String event;

    Output(String event) {
        this.event = event;
    }

    @Override
    public String getEvent() {
        return this.toString();
    }

    @Override
    public String getEventName() {
        return event;
    }
}
