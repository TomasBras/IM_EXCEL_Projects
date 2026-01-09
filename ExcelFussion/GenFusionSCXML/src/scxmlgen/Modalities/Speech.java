package scxmlgen.Modalities;

import scxmlgen.interfaces.IModality;

public enum Speech implements IModality {


    CALCULAR_MEDIA("[SPEECH][CALCULAR_MEDIA]", 5000),
    DESTACAR_APROVADOS_REPROVADOS("[SPEECH][DESTACAR_APROVADOS_REPROVADOS]", 5000),
    INSERIR_COLUNAS("[SPEECH][INSERIR_COLUNAS]", 5000),
    MELHORIA_REAL("[SPEECH][MELHORIA_REAL]", 5000),
    MELHORIA_POSSIVEL("[SPEECH][MELHORIA_POSSIVEL]", 5000),

    GERAR_GRAFICO_TURMA("[SPEECH][GERAR_GRAFICO_TURMA]", 5000),
    GERAR_GRAFICO_BARRAS_ALUNO("[SPEECH][GERAR_GRAFICO_BARRAS_ALUNO]", 5000),
    GERAR_GRAFICO_PERGUNTAS_T2("[SPEECH][GERAR_GRAFICO_PERGUNTAS_T2]", 5000),


    CRIAR_PIVOT_TABLE("[SPEECH][CRIAR_PIVOT_TABLE]", 5000),
    OPERACOES_MATEMATICAS("[SPEECH][OPERACOES_MATEMATICAS]", 5000),

   
    APAGAR_TODOS_GRAFICOS("[SPEECH][APAGAR_TODOS_GRAFICOS]", 5000),
    ATUALIZAR_NOTAS("[SPEECH][ATUALIZAR_NOTAS]", 5000),
    GUARDAR_FICHEIRO("[SPEECH][GUARDAR_FICHEIRO]", 5000),

    CONFIRMAR("[SPEECH][CONFIRMAR]", 5000),
    CANCELAR("[SPEECH][CANCELAR]", 5000),

  
    CLOSE_EXCEL("[SPEECH][CLOSE_EXCEL]", 5000),
    HELPER("[SPEECH][HELPER]", 5000);

    private final String event;
    private final int timeout;

    Speech(String event, int timeout) {
        this.event = event;
        this.timeout = timeout;
    }

    @Override
    public int getTimeOut() {
        return timeout;
    }

    @Override
    public String getEventName() {
        return event;
    }

    @Override
    public String getEvName() {
        return getModalityName().toLowerCase() + event.toLowerCase();
    }
}
