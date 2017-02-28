package cipry.ro.hs2015.pedometer;

import android.os.Bundle;
import android.preference.PreferenceActivity;

import cipry.ro.hs2015.R;

public class Settings extends PreferenceActivity {
    
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        
        addPreferencesFromResource(R.xml.preferences);
    }
}
