package cipry.ro.hs2015.pedometer;

import android.app.Service;
import android.speech.tts.TextToSpeech;
import android.text.format.Time;
import android.util.Log;

import java.util.Locale;

public class Utils implements TextToSpeech.OnInitListener {
    private static final String TAG = "Utils";
    private Service mService;

    private static Utils instance = null;

    private Utils() {
    }
     
    public static Utils getInstance() {
        if (instance == null) {
            instance = new Utils();
        }
        return instance;
    }
    
    public void setService(Service service) {
        mService = service;
    }
    
    
    private TextToSpeech mTts;
    private boolean mSpeak = false;
    private boolean mSpeakingEngineAvailable = false;

    public void initTTS() {
        Log.i(TAG, "Initializing TextToSpeech...");
        mTts = new TextToSpeech(mService,
            this
            );
    }
    public void shutdownTTS() {
        Log.i(TAG, "Shutting Down TextToSpeech...");

        mSpeakingEngineAvailable = false;
        mTts.shutdown();
        Log.i(TAG, "TextToSpeech Shut Down.");

    }
    public void say(String text) {
        if (mSpeak && mSpeakingEngineAvailable) {
            mTts.speak(text,
                    TextToSpeech.QUEUE_ADD,
                    null);
        }
    }

    
    public void onInit(int status) {
        if (status == TextToSpeech.SUCCESS) {
            int result = mTts.setLanguage(Locale.US);
            if (result == TextToSpeech.LANG_MISSING_DATA ||
                result == TextToSpeech.LANG_NOT_SUPPORTED) {
               // Language data is missing or the language is not supported.
                Log.e(TAG, "Language is not available.");
            } else {
                Log.i(TAG, "TextToSpeech Initialized.");
                mSpeakingEngineAvailable = true;
            }
        } else {
            // Initialization failed.
            Log.e(TAG, "Could not initialize TextToSpeech.");
        }
    }

    public void setSpeak(boolean speak) {
        mSpeak = speak;
    }

    public boolean isSpeakingEnabled() {
        return mSpeak;
    }

    public boolean isSpeakingNow() {
        return mTts.isSpeaking();
    }

    public void ding() {
    }
        
    public static long currentTimeInMillis() {
        Time time = new Time();
        time.setToNow();
        return time.toMillis(false);
    }
}
