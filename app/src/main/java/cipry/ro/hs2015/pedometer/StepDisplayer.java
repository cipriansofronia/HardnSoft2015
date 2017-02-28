package cipry.ro.hs2015.pedometer;

import java.util.ArrayList;

public class StepDisplayer implements StepListener, SpeakingTimer.Listener {

    private int mCount = 0;
    PedometerSettings mSettings;
    Utils mUtils;

    public StepDisplayer(PedometerSettings settings, Utils utils) {
        mUtils = utils;
        mSettings = settings;
        notifyListener();
    }
    public void setUtils(Utils utils) {
        mUtils = utils;
    }

    public void setSteps(int steps) {
        mCount = steps;
        notifyListener();
    }
    public void onStep() {
        mCount ++;
        notifyListener();
    }
    public void reloadSettings() {
        notifyListener();
    }
    public void passValue() {
    }
    
    
    public interface Listener {
        public void stepsChanged(int value);
        public void passValue();
    }
    private ArrayList<Listener> mListeners = new ArrayList<Listener>();

    public void addListener(Listener l) {
        mListeners.add(l);
    }
    public void notifyListener() {
        for (Listener listener : mListeners) {
            listener.stepsChanged((int)mCount);
        }
    }

    
    public void speak() {
        if (mSettings.shouldTellSteps()) { 
            if (mCount > 0) {
                mUtils.say("" + mCount + " steps");
            }
        }
    }
    
    
}
