package scxmlgen.Modalities;

import scxmlgen.interfaces.IModality;

public enum Gestures implements IModality {

    SWIPELEFT("[GESTURES][SWIPELEFT]", 1500),
    SWIPERIGHT("[GESTURES][SWIPERIGHT]", 1500),
    SWIPEUP("[GESTURES][SWIPEUP]", 1500),
    SWIPEDOWN("[GESTURES][SWIPEDOWN]", 1500),


    HANDGRAB("[GESTURES][HANDGRAB]", 5000),

    CALCULATEAVERAGE("[GESTURES][CALCULATEAVERAGE]", 3000),
    INSERTCOLUMN("[GESTURES][INSERTCOLUMN]", 3000),
    STUDENTSAPPROVED("[GESTURES][STUDENTSAPPROVED]", 3000),
    STUDENTSFAILED("[GESTURES][STUDENTSFAILED]", 3000),


    UNDOLASTACTION("[GESTURES][UNDOLASTACTION]", 3000),


    ZOOMIN("[GESTURES][ZOOMIN]", 1500),
    ZOOMOUT("[GESTURES][ZOOMOUT]", 1500);

    private final String event;
    private final int timeout;

    Gestures(String event, int timeout) {
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
