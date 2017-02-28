package cipry.ro.hs2015;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.app.ProgressDialog;
import android.bluetooth.BluetoothAdapter;
import android.bluetooth.BluetoothDevice;
import android.bluetooth.BluetoothSocket;
import android.content.ComponentName;
import android.content.Context;
import android.content.Intent;
import android.content.ServiceConnection;
import android.content.SharedPreferences;
import android.graphics.Typeface;
import android.location.Location;
import android.location.LocationListener;
import android.location.LocationManager;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Environment;
import android.os.Handler;
import android.os.IBinder;
import android.os.Looper;
import android.os.Message;
import android.os.StrictMode;
import android.preference.PreferenceManager;
import android.text.TextUtils;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.WindowManager;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.widget.Toast;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Random;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import cipry.ro.hs2015.frequency.Audio;
import cipry.ro.hs2015.pedometer.PedometerSettings;
import cipry.ro.hs2015.pedometer.Settings;
import cipry.ro.hs2015.pedometer.StepService;
import cipry.ro.hs2015.pedometer.Utils;

import cipry.ro.hs2015.twitter.WebViewActivity;
import twitter4j.StatusUpdate;
import twitter4j.Twitter;
import twitter4j.TwitterException;
import twitter4j.TwitterFactory;
import twitter4j.User;
import twitter4j.auth.AccessToken;
import twitter4j.auth.RequestToken;
import twitter4j.conf.Configuration;
import twitter4j.conf.ConfigurationBuilder;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class MainActivity extends Activity {

    /* Shared preference keys */
    private static final String PREF_NAME = "sample_twitter_pref";
    private static final String PREF_KEY_OAUTH_TOKEN = "oauth_token";
    private static final String PREF_KEY_OAUTH_SECRET = "oauth_token_secret";
    private static final String PREF_KEY_TWITTER_LOGIN = "is_twitter_loggedin";
    private static final String PREF_USER_NAME = "twitter_user_name";

    /* Any number for uniquely distinguish your request */
    public static final int WEBVIEW_REQUEST_CODE = 100;

    private static Twitter twitter;
    private static RequestToken requestToken;

    private static SharedPreferences mSharedPreferences;

    private String consumerKey = null;
    private String consumerSecret = null;
    private String callbackUrl = null;
    private String oAuthVerifier = null;

    private static final String TAG = "Pedometer";
    private SharedPreferences mSettings;
    private PedometerSettings mPedometerSettings;
    private Utils mUtils;

    private TextView mStepValueView;
    private TextView mPaceValueView;
    private TextView mDistanceValueView, mFrequencyView;
    private TextView mSpeedValueView, titleHS, frequencyView;
    private TextView mCaloriesValueView;
    TextView mDesiredPaceView;
    private int mStepValue;
    private int mPaceValue;
    private float mDistanceValue;
    private float mSpeedValue;
    private int mCaloriesValue;
    private float mDesiredPaceOrSpeed;
    private int mMaintain;
    private boolean mIsMetric;
    private float mMaintainInc;
    private boolean mQuitting = false; // Set when user selected Quit from menu, can be used by onPause, onStop, onDestroy
    private Handler mHandler2;
    private Audio mAudio;
    private boolean isLoggedIn;
    private ScheduledExecutorService scheduleTaskExecutor;

    // private final static int INTERVAL = 1000 * 60;
    // private Handler mHandlerTwitts;

    private BluetoothAdapter mBluetoothAdapter;
    private BluetoothSocket mmSocket;
    private BluetoothDevice mmDevice;
    private OutputStream mmOutputStream;
    private InputStream mmInputStream;
    private String receivedString = "0";
    private byte[] readBuffer;
    private int readBufferPosition;
    private Boolean start_enabled = false, connectedToDevice = false, stopWorker = true;

    private int heartRateGlobal=0, bodyTempGlobal=0, airTempGlobal=0, humidityGlobal=0, ambiLightGlobal=0,
            airTempTreasureGlobal=0, magneticTreasureGlobal=0, blinkRateGlobal=0;
    private ArrayList<String> rfidList = new ArrayList<>();
    private String rfidGlobal = "none", remoteTreasureGlobal = "";
    private int wayPoint = 0;

    private Location myLocation;
    private double frequencyGlobal;

    private String hexIR[] = {
        //new remote codes
        "17B4A228", "530DB67C", //POWER button - 0,1
        "C1D0902A", "55EF31B6", //0 - 2,3
        "C2D091BF", "56EF334B", //1 - 4,5
        "528A5222", "17313DCE", //2 - 6,7
        "E4400B14", "50216988", //3 - 8,9
        "6D6E8E28", "321579D4", //4 - :,;
        "6E6E8FBB", "33157B67", //5 - <,=
        "DF75FD61", "A41CE90D", //6 - >,?
        "FE084450", "69E9A2C4", //7 - @, A
        "1DDBEF8C", "593503E0", //8 - B,C
        "1ADBEAD5", "5634FF29", //9 - D,E
        "1AB4A6E1", "560DBB35", //mute button - F,G
        "D1D4CC60", "967BB80C", //channel + - T,U
        "CED4C7A9", "937BB355", //channel - - W,V
        "92F27C7C", "411C1BB0", //-/-- - L, M
        "D0264A08", "94CD35B4", //AV - N,O
        "940EAE71", "58B59A1D", //volume down - P, Q
        "970EB328", "5BB59ED4", //volume up - Z, [
        "D1D4CC60", "967BB80C", //up arrow -
        "937BB355", "CED4C7A9", //down arrow -
        "5BB59ED4", "970EB328", //right arrow -
        "970EB328", "5BB59ED4", //left arrow -
        "F47944C5", "A2A2E3F9", //ok button - \, ]
        "903F952A", "54E680D6", //menu button - ^,_
        "C50C30BA", "89B31C66", //teletext red - `,a
        "F53B40EB", "4711A1B7", //teletext green - b,c
        "3A872A2C", "8C5D8AF8", //teletext yellow - d,e
        "210F274C", "CF38C680", //teletext blue - f,g
        "E824E15C", "ACCBCD08", //menu 1 button - h,i
        "CBE66003", "1DBCC0CF", //menu x2 button - j,k
        "5AA370F2", "8CD1026",  //menu x3 button - l,m
        "31A318D4", "9D847748"  //menu x4 button - n,o
    }; // IR hex codes is an array of pointers to char

    /**
     * True, when service is running.
     */
    private boolean mIsRunning;

    /** Called when the activity is first created. */
    @Override
    public void onCreate(Bundle savedInstanceState) {
        Log.i(TAG, "[ACTIVITY] onCreate");
        super.onCreate(savedInstanceState);

        setContentView(R.layout.main);

        getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);

        mStepValue = 0;
        mPaceValue = 0;

        turnOnBT();

        mUtils = Utils.getInstance();

        /* initializing twitter parameters from string.xml */
        initTwitterConfigs();

        /* Enabling strict mode */
        StrictMode.ThreadPolicy policy = new StrictMode.ThreadPolicy.Builder().permitAll().build();
        StrictMode.setThreadPolicy(policy);

        /* Check if required twitter keys are set */
        if (TextUtils.isEmpty(consumerKey) || TextUtils.isEmpty(consumerSecret)) {
            Toast.makeText(this, "Twitter key and secret not configured",
                    Toast.LENGTH_SHORT).show();
            return;
        }

        /* Initialize application preferences */
        mSharedPreferences = getSharedPreferences(PREF_NAME, 0);

        isLoggedIn = mSharedPreferences.getBoolean(PREF_KEY_TWITTER_LOGIN, false);

        if (!isLoggedIn) {
            Uri uri = getIntent().getData();

            if (uri != null && uri.toString().startsWith(callbackUrl)) {

                String verifier = uri.getQueryParameter(oAuthVerifier);

                try {

					/* Getting oAuth authentication token */
                    AccessToken accessToken = twitter.getOAuthAccessToken(requestToken, verifier);

					/* Getting user id form access token */
                    long userID = accessToken.getUserId();
                    final User user = twitter.showUser(userID);
                    final String username = user.getName();

					/* save updated token */
                    saveTwitterInfo(accessToken);

                    /*loginLayout.setVisibility(View.GONE);
                    shareLayout.setVisibility(View.VISIBLE);
                    userName.setText(getString(R.string.hello) + username);*/

                } catch (Exception e) {
                    Log.e("Failed to login Twitter!!", e.getMessage());
                }
            }
        }

        // Acquire a reference to the system Location Manager
        LocationManager locationManager = (LocationManager) this.getSystemService(Context.LOCATION_SERVICE);

        // Define a listener that responds to location updates
        LocationListener locationListener = new LocationListener() {
            public void onLocationChanged(Location location) {
                // Called when a new location is found by the network location provider.
                //Log.e("", "" + location);
                myLocation = location;
            }

            public void onStatusChanged(String provider, int status, Bundle extras) {}

            public void onProviderEnabled(String provider) {}

            public void onProviderDisabled(String provider) {}
        };

        // Register the listener with the Location Manager to receive location updates
        locationManager.requestLocationUpdates(LocationManager.NETWORK_PROVIDER, 0, 0, locationListener);

        Typeface myTypeface = Typeface.createFromAsset(getAssets(), "fonts/orator_std.otf");
        titleHS = (TextView) findViewById(R.id.title_HS);
        titleHS.setTypeface(myTypeface);

        frequencyView = (TextView) findViewById(R.id.desired_frequency_label);
        frequencyView.setTypeface(myTypeface);

        Button saveBtn = (Button) findViewById(R.id.idSave);
        saveBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (start_enabled) {
                    switch (wayPoint) {
                        case 1: if (saveExcelStart(getApplicationContext(), "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 2: if (readWriteStation(getApplicationContext(), 1, "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 3: if (readWriteStation(getApplicationContext(), 2, "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 4: if (readWriteStation(getApplicationContext(), 3, "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 5: if (readWriteStation(getApplicationContext(), 4, "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 6: if (readWriteStation(getApplicationContext(), 5, "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        case 7: if (saveExcelFinish(getApplicationContext(), "Suceava2HS2015.xls"))
                                    Toast.makeText(getApplicationContext(), "Saved!", Toast.LENGTH_SHORT).show();
                                else
                                    Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                                break;
                        default:
                            Toast.makeText(getApplicationContext(), "Not saved!", Toast.LENGTH_SHORT).show();
                            break;
                    }
                } else {
                    Toast.makeText(getApplicationContext(), "Not connected!", Toast.LENGTH_SHORT).show();
                }
            }
        });

    }

    private void turnOnBT() {
        mBluetoothAdapter = BluetoothAdapter.getDefaultAdapter();
        if (mBluetoothAdapter == null) {
            Toast.makeText(getBaseContext(), "No bluetooth adapter available!", Toast.LENGTH_SHORT).show();
            //myLabel.setText("No bluetooth adapter available");
            return;
        }

        if(!mBluetoothAdapter.isEnabled()) {
            if (mBluetoothAdapter.getState() == BluetoothAdapter.STATE_OFF) {
                mBluetoothAdapter.enable();
                Toast.makeText(getBaseContext(), "Bluetooth off! Enabling now...", Toast.LENGTH_SHORT).show();
            } else {
                //State.INTERMEDIATE_STATE;
            }
        }
    }

    private void connectBTDevice() throws IOException {

        if (mBluetoothAdapter.isEnabled()) {

            Set<BluetoothDevice> pairedDevices = mBluetoothAdapter.getBondedDevices();

            if (pairedDevices.size() > 0) {
                for (BluetoothDevice device : pairedDevices) {
                    if (device.getName().equals("HC-06")) {
                        mmDevice = device;
                        //Toast.makeText(getBaseContext(), "linvor Found! NOT Connected!", Toast.LENGTH_SHORT).show();
                        break;
                    }
                }
            }

            if(mmDevice == null){
                if (true) Log.e("Error", "Device Negasit!");
                return;
            } else {
                UUID uuid = UUID.fromString("00001101-0000-1000-8000-00805F9B34FB"); // Standard SerialPortService ID

                try {
                    mmSocket = mmDevice.createRfcommSocketToServiceRecord(uuid);
                } catch (IOException e) {
                    Log.d("ERR", "In onResume() and socket create failed: " + e.getMessage());
                    return;
                }
                mBluetoothAdapter.cancelDiscovery();

                try {
                    mmSocket.connect();
                } catch (IOException e) {
                    e.printStackTrace();
                    runOnUiThread(new Runnable() {
                        public void run() {
                            Toast.makeText(getBaseContext(), "Connection Failed!", Toast.LENGTH_SHORT).show();
                        }
                    });
                    try {
                        disconnectBTDevice();
                    } catch (IOException e2) {
                        Log.d("ERR", "Unable to close socket during connection failure " + e2.getMessage());
                        return;
                    }
                    return;
                }

                if (mmSocket.isConnected()) {
                    try {
                        mmOutputStream = mmSocket.getOutputStream();
                        mmInputStream = mmSocket.getInputStream();
                        connectedToDevice = true;
                    } catch (IOException e) {
                        Log.d("ERR", "In onResume() and output stream creation failed:" + e.getMessage());
                        connectedToDevice = false;
                        return;
                    }

                    /*Handler handler = new Handler(Looper.getMainLooper());
                    handler.postDelayed(new Runnable() {
                        @Override
                        public void run() {
                        }
                    }, 1);*/
                    // listenAndEvolve();

                    runOnUiThread(new Runnable() {
                        public void run() {
                            Toast.makeText(getBaseContext(), "Bluetooth Connected!", Toast.LENGTH_SHORT).show();
                        }
                    });
                } else {
                    Log.e("Error","Connection Failed!!");
                    runOnUiThread(new Runnable() {
                        public void run() {
                            Toast.makeText(getBaseContext(), "Connection Failed!", Toast.LENGTH_SHORT).show();
                        }
                    });
                    connectedToDevice = false;
                    disconnectBTDevice();
                    return;
                }
            }
        }
    }

    private void listenAndEvolve() {
        final Handler handler = new Handler();
        final byte delimiter = 10; // This is the ASCII code for a newline character

        stopWorker = false;
        readBufferPosition = 0;
        readBuffer = new byte[1024];
        Thread workerThread = new Thread(new Runnable() {
            @SuppressLint("NewApi")
            public void run() {
                while (!Thread.currentThread().isInterrupted() && !stopWorker) {
                    try {
                        if (mmSocket.isConnected()) {

                            connectedToDevice = true;

                            int bytesAvailable = mmInputStream.available();
                            if (bytesAvailable > 0) {
                                byte[] packetBytes = new byte[bytesAvailable];
                                mmInputStream.read(packetBytes);
                                for (int i = 0; i < bytesAvailable; i++) {
                                    byte b = packetBytes[i];
                                    if (b == delimiter) {
                                        byte[] encodedBytes = new byte[readBufferPosition];
                                        System.arraycopy(readBuffer, 0,	encodedBytes, 0, encodedBytes.length);
                                        final String data = new String(encodedBytes, "US-ASCII");
                                        readBufferPosition = 0;

                                        handler.post(new Runnable() {
                                            public void run() {
                                                receivedString = data;
                                                //Log.d("<-- received data: ", String.valueOf(receivedString));

                                                int heartRate=0, bodyTemp=0, airTemp=0, humidity=0, ambiLight=0, remoteTreasure=70,
                                                        airTempTreasure=0, magneticTreasure=0, blinkRate=0;

                                                char[] buffer = toCharacterArray(receivedString);
                                                int index = 0;
                                                int aux = 0, number = 0;
                                                int length = receivedString.length();
                                                for(int i=0; i<length; i++)
                                                {
                                                    if (buffer[i] != '#') {
                                                        if(buffer[i] != '@'){
                                                            aux=aux*10+(buffer[i]-'0');
                                                        }
                                                        else
                                                        {
                                                            number = aux;
                                                            switch(index) {
                                                                case 0: heartRate = number; break;
                                                                case 1: bodyTemp = number; break;
                                                                case 2: airTemp = number; break;
                                                                case 3: humidity = number; break;
                                                                case 4: ambiLight = number; break;
                                                                case 5: remoteTreasure = number; break;
                                                                case 6: airTempTreasure = number; break;
                                                                case 7: magneticTreasure = number; break;
                                                                case 8: blinkRate = number; break;
                                                            }
                                                            index++; aux = 0;
                                                        }
                                                    }
                                                }

                                                /*Log.e("Received: ", heartRate + ", " + bodyTemp + ", " + airTemp + ", " + humidity +
                                                        ", " + ambiLight + ", " + remoteTreasure + ", " + airTempTreasure
                                                        + ", " + magneticTreasure + ", " + blinkRate);*/

                                                if (receivedString.contains("#")) {
                                                    String rfid[] = receivedString.split("#");
                                                    rfidGlobal = rfid[1];

                                                    if (!rfidGlobal.contains("null") && (rfidGlobal.length() <= 9)) {
                                                        if (rfidList.isEmpty()) {
                                                            rfidList.add(rfidGlobal);
                                                            startGame();
                                                            wayPoint = rfidList.size();

                                                            runOnUiThread(new Runnable() {
                                                                public void run() {
                                                                    mFrequencyView.setText("WayPoint: " + rfidGlobal + "(" + wayPoint + ")");
                                                                }
                                                            });
                                                        } else if (!rfidList.contains(rfidGlobal)) {
                                                            rfidList.add(rfidGlobal);
                                                            wayPoint = rfidList.size();
                                                            remoteTreasureGlobal = "";
                                                            airTempTreasureGlobal = 0;
                                                            magneticTreasureGlobal = 0;
                                                            blinkRateGlobal = 0;

                                                            runOnUiThread(new Runnable() {
                                                                public void run() {
                                                                    mFrequencyView.setText("WayPoint: " + rfidGlobal + "(" + wayPoint + ")");
                                                                }
                                                            });
                                                            // deja next way point
                                                        }
                                                    }
                                                    //Log.e("RFID: ", rfid[1]);
                                                }

                                                heartRateGlobal = heartRate;
                                                bodyTempGlobal = bodyTemp;
                                                airTempGlobal = airTemp;
                                                humidityGlobal = humidity;
                                                ambiLightGlobal = ambiLight;

                                                if (remoteTreasure != 70 && remoteTreasure >= 0 && remoteTreasure <= 64) {
                                                    remoteTreasureGlobal = hexIR[remoteTreasure];
                                                }
                                                if (airTempTreasure > 35) {
                                                    airTempTreasureGlobal = airTempTreasure;
                                                }
                                                if (magneticTreasure != 0 && (magneticTreasure > 300 || magneticTreasure < 200) ) {
                                                    // Y = (X-A)/(B-A) * (D-C) + C
                                                    //int A = 1, B = 1024, C = -650, D = 650;
                                                    //magneticTreasureGlobal = (magneticTreasure-A)/(B-A) * (D-C) + C;
                                                    magneticTreasureGlobal = 1;
                                                }
                                                if (blinkRate != 0 && blinkRate > 4) {
                                                    blinkRateGlobal = blinkRate;
                                                }

                                                int a[] = {0,0,0,0,0};
                                                if (remoteTreasureGlobal.length() > 1) { a[0]=1; }
                                                if (airTempTreasureGlobal > 35) { a[1]=1; }
                                                if (magneticTreasureGlobal == 1) { a[2]=1; }
                                                if (blinkRateGlobal != 0) { a[3]=1; }
                                                if (frequencyGlobal != 0) { a[4]=1; }

                                                final StringBuilder str = new StringBuilder();
                                                str.append("Treasures found: ");
                                                if (remoteTreasureGlobal.length() > 1) {
                                                    str.append(" Remote: ").append(remoteTreasureGlobal);
                                                }
                                                if (airTempTreasureGlobal > 35) {
                                                    str.append(" Temp: ").append(airTempTreasureGlobal);
                                                }
                                                if (magneticTreasureGlobal == 1){
                                                    str.append(" magnetic filed found ");
                                                }
                                                if (blinkRateGlobal != 0) {
                                                    str.append(" blink rate: ").append(blinkRateGlobal);
                                                }
                                                if (frequencyGlobal != 0) {
                                                    str.append(" Frequency: ").append((int)frequencyGlobal);
                                                }
                                                if (a[0] == 0 && a[1] == 0 && a[2] == 0 && a[3] == 0 && a[4] == 0) {
                                                    str.append(" none ");
                                                }

                                                runOnUiThread(new Runnable() {
                                                    public void run() {
                                                        mFrequencyView.setText("WayPoint: " + rfidGlobal + "(" + wayPoint + ") " + str.toString());
                                                    }
                                                });

                                            }
                                        });
                                    } else {
                                        readBuffer[readBufferPosition++] = b;
                                    }
                                }
                            } else {
                                // ? nu se primeste caractere
                                // if (true) Log.e("Error!"," Deconectat de la BT! ");
                                // if mmSocket.isConnected()
                                //runningConditions("Not receiving data!");
                            }
                        } else {
                            if (true) Log.e("Error!"," Deconectat de la BT! ");
                            connectedToDevice = false;
                            // initialState("Disconnected!");
                        }
                    } catch (IOException ex) {
                        stopWorker = true;
                        connectedToDevice = false;
                        Log.e("Error!","" + ex.getMessage());
                    }
                }
            }
        });

        workerThread.start();
    }

    private char[] toCharacterArray(String s) {
        if (s == null) {
            return null;
        }
        char[] c = new char[s.length()];
        for (int i = 0; i < s.length(); i++) {
            c[i] = s.charAt(i);
        }

        return c;
    }

    private void disconnectBTDevice() throws IOException {
        stopWorker = true;

        Log.d("TAG", "Disconnecting from device...");
        if (mmOutputStream != null) {
            try {
                mmOutputStream.close();
                mmInputStream.close();
                connectedToDevice = false;
            } catch (IOException e) {
                Log.d("TAG", "In onPause() and failed to flush output stream: "	+ e.getMessage());
                runOnUiThread(new Runnable() {
                    public void run() {
                        Toast.makeText(getBaseContext(), "Socket failed", Toast.LENGTH_SHORT).show();
                    }
                });

                return;
            }
        }

        if (mmSocket != null) {
            try {
                mmSocket.close();
            } catch (IOException e2) {
                Log.d("TAG", "In onPause() and failed to close socket." + e2.getMessage());
                runOnUiThread(new Runnable() {
                    public void run() {
                        Toast.makeText(getBaseContext(), "Socket failed", Toast.LENGTH_SHORT).show();
                    }
                });
                return;
            }
        }

        runOnUiThread(new Runnable() {
            public void run() {
                Toast.makeText(getBaseContext(), "Disconnected from device!", Toast.LENGTH_SHORT).show();
            }
        });

    }

    @Override
    public void onBackPressed() {
        //super.onBackPressed();
        new Handler().postDelayed(new Runnable() {
            public void run() {
                openOptionsMenu();
            }
        }, 5);
    }

    @Override
    protected void onStart() {
        Log.i(TAG, "[ACTIVITY] onStart");
        super.onStart();
    }

    @Override
    protected void onResume() {
        Log.i(TAG, "[ACTIVITY] onResume");
        super.onResume();

        initHandler();
        initSound();

        if (mIsRunning) mAudio.startRecording(0);

        mSettings = PreferenceManager.getDefaultSharedPreferences(this);
        mPedometerSettings = new PedometerSettings(mSettings);

        mUtils.setSpeak(mSettings.getBoolean("speak", false));

        // Read from preferences if the service was running on the last onPause
        mIsRunning = mPedometerSettings.isServiceRunning();

        // Start the service if this is considered to be an application start (last onPause was long ago)
        /*if (!mIsRunning && mPedometerSettings.isNewStart()) {
            startStepService();
            bindStepService();
        }
        else */if (mIsRunning) {
            bindStepService();
        }

        mPedometerSettings.clearServiceRunning();

        mStepValueView     = (TextView) findViewById(R.id.step_value);
        mPaceValueView     = (TextView) findViewById(R.id.pace_value);
        mDistanceValueView = (TextView) findViewById(R.id.distance_value);
        mSpeedValueView    = (TextView) findViewById(R.id.speed_value);
        mCaloriesValueView = (TextView) findViewById(R.id.calories_value);
        mDesiredPaceView   = (TextView) findViewById(R.id.desired_pace_value);

        mIsMetric = mPedometerSettings.isMetric();
        ((TextView) findViewById(R.id.distance_units)).setText(getString(
                mIsMetric
                        ? R.string.kilometers
                        : R.string.miles
        ));
        ((TextView) findViewById(R.id.speed_units)).setText(getString(
                mIsMetric
                        ? R.string.kilometers_per_hour
                        : R.string.miles_per_hour
        ));

        mMaintain = mPedometerSettings.getMaintainOption();
        ((LinearLayout) this.findViewById(R.id.desired_pace_control)).setVisibility(
                mMaintain != PedometerSettings.M_NONE
                        ? View.VISIBLE
                        : View.GONE
        );
        if (mMaintain == PedometerSettings.M_PACE) {
            mMaintainInc = 5f;
            mDesiredPaceOrSpeed = (float)mPedometerSettings.getDesiredPace();
        }
        else
        if (mMaintain == PedometerSettings.M_SPEED) {
            mDesiredPaceOrSpeed = mPedometerSettings.getDesiredSpeed();
            mMaintainInc = 0.1f;
        }
        Button button1 = (Button) findViewById(R.id.button_desired_pace_lower);
        button1.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                mDesiredPaceOrSpeed -= mMaintainInc;
                mDesiredPaceOrSpeed = Math.round(mDesiredPaceOrSpeed * 10) / 10f;
                displayDesiredPaceOrSpeed();
                setDesiredPaceOrSpeed(mDesiredPaceOrSpeed);

                readWriteStation(getApplicationContext(), 1, "Suceava2HS2015.xls");
                saveExcelFinish(getApplicationContext(), "Suceava2HS2015.xls");

            }
        });
        Button button2 = (Button) findViewById(R.id.button_desired_pace_raise);
        button2.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                mDesiredPaceOrSpeed += mMaintainInc;
                mDesiredPaceOrSpeed = Math.round(mDesiredPaceOrSpeed * 10) / 10f;
                displayDesiredPaceOrSpeed();
                setDesiredPaceOrSpeed(mDesiredPaceOrSpeed);

                saveExcelStart(getApplicationContext(), "Suceava2HS2015.xls");


            }
        });
        if (mMaintain != PedometerSettings.M_NONE) {
            ((TextView) findViewById(R.id.desired_pace_label)).setText(
                    mMaintain == PedometerSettings.M_PACE
                            ? R.string.desired_pace
                            : R.string.desired_speed
            );
        }


        displayDesiredPaceOrSpeed();
    }

    private void displayDesiredPaceOrSpeed() {
        if (mMaintain == PedometerSettings.M_PACE) {
            mDesiredPaceView.setText("" + (int)mDesiredPaceOrSpeed);
        }
        else {
            mDesiredPaceView.setText("" + mDesiredPaceOrSpeed);
        }
    }

    public void initHandler(){
        mHandler2 = new Handler(Looper.getMainLooper()){
            public void handleMessage(Message message) {
                Bundle bundle = message.getData();
                double frequency = bundle.getDouble("frequency");
                double dB = bundle.getDouble("dB");
                // Log.d("MESSAGE",""+frequency+" "+dB);
                frequencyGlobal = frequency;
                // showResult(frequency, dB);
            }
        };
    }

    public void initSound(){
        mFrequencyView     = (TextView) findViewById(R.id.desired_frequency_label);
        mFrequencyView.setText("Data to be displayed!");
        mAudio = new Audio(mHandler2);
    }

    public void showResult(double frequency, double dB){
        if (true/*APP_MODE == 0*/){
            if (!mIsRunning){
                mFrequencyView.setText("Frequency: 0 Hz");
            }
            else{
                frequency = Math.round(frequency);
                mFrequencyView.setText("Frequency: " + (int)frequency + " Hz");
            }
        }
    }

    @Override
    protected void onPause() {
        Log.i(TAG, "[ACTIVITY] onPause");
        if (mIsRunning) {
            unbindStepService();
        }
        if (mQuitting) {
            mPedometerSettings.saveServiceRunningWithNullTimestamp(mIsRunning);
        }
        else {
            mPedometerSettings.saveServiceRunningWithTimestamp(mIsRunning);
        }

        super.onPause();
        savePaceSetting();
        mAudio.stopRecording();

    }

    @Override
    protected void onStop() {
        Log.i(TAG, "[ACTIVITY] onStop");
        super.onStop();
        mAudio.release();

    }

    protected void onDestroy() {
        Log.i(TAG, "[ACTIVITY] onDestroy");
        super.onDestroy();
    }

    protected void onRestart() {
        Log.i(TAG, "[ACTIVITY] onRestart");
        super.onDestroy();
    }

    private void setDesiredPaceOrSpeed(float desiredPaceOrSpeed) {
        if (mService != null) {
            if (mMaintain == PedometerSettings.M_PACE) {
                mService.setDesiredPace((int)desiredPaceOrSpeed);
            }
            else
            if (mMaintain == PedometerSettings.M_SPEED) {
                mService.setDesiredSpeed(desiredPaceOrSpeed);
            }
        }
    }

    private void savePaceSetting() {
        mPedometerSettings.savePaceOrSpeedSetting(mMaintain, mDesiredPaceOrSpeed);
    }

    private StepService mService;

    private ServiceConnection mConnection = new ServiceConnection() {
        public void onServiceConnected(ComponentName className, IBinder service) {
            mService = ((StepService.StepBinder)service).getService();

            mService.registerCallback(mCallback);
            mService.reloadSettings();

        }

        public void onServiceDisconnected(ComponentName className) {
            mService = null;
        }
    };


    private void startStepService() {
        if (! mIsRunning) {
            Log.i(TAG, "[SERVICE] Start");
            mIsRunning = true;
            startService(new Intent(MainActivity.this,
                    StepService.class));
        }
    }

    private void bindStepService() {
        Log.i(TAG, "[SERVICE] Bind");
        bindService(new Intent(MainActivity.this,
                StepService.class), mConnection, Context.BIND_AUTO_CREATE + Context.BIND_DEBUG_UNBIND);
    }

    private void unbindStepService() {
        Log.i(TAG, "[SERVICE] Unbind");
        if (mIsRunning) unbindService(mConnection);
    }

    private void stopStepService() {
        Log.i(TAG, "[SERVICE] Stop");
        if (mService != null) {
            Log.i(TAG, "[SERVICE] stopService");
            stopService(new Intent(MainActivity.this,
                    StepService.class));
        }
        mIsRunning = false;
    }

    private void resetValues(boolean updateDisplay) {
        if (mService != null && mIsRunning) {
            mService.resetValues();
        }
        else {
            mStepValueView.setText("0");
            mPaceValueView.setText("0");
            mDistanceValueView.setText("0");
            mSpeedValueView.setText("0");
            mCaloriesValueView.setText("0");
            SharedPreferences state = getSharedPreferences("state", 0);
            SharedPreferences.Editor stateEditor = state.edit();
            if (updateDisplay) {
                stateEditor.putInt("steps", 0);
                stateEditor.putInt("pace", 0);
                stateEditor.putFloat("distance", 0);
                stateEditor.putFloat("speed", 0);
                stateEditor.putFloat("calories", 0);
                stateEditor.commit();
            }
        }
    }

    private static final int MENU_SETTINGS = 8;
    private static final int MENU_QUIT     = 9;

    private static final int MENU_PAUSE = 1;
    private static final int MENU_RESUME = 2;
    private static final int MENU_RESET = 3;
    private static final int MENU_LOGIN = 4;
    private static final int MENU_CONBT = 5;
    private static final int MENU_DISCONBT = 6;

    /* Creates the menu items */
    public boolean onPrepareOptionsMenu(Menu menu) {
        menu.clear();
        menu.add(0, MENU_CONBT, 0, R.string.connect)
                .setShortcut('1', 'c');
        menu.add(0, MENU_DISCONBT, 0, R.string.disconnect)
                .setShortcut('1', 'd');
        if (mIsRunning) {
            menu.add(0, MENU_PAUSE, 0, R.string.pause)
                    .setIcon(android.R.drawable.ic_media_pause)
                    .setShortcut('1', 'p');
        }
        else {
            menu.add(0, MENU_RESUME, 0, R.string.resume)
                    .setIcon(android.R.drawable.ic_media_play)
                    .setShortcut('1', 'p');
        }
        menu.add(0, MENU_RESET, 0, R.string.reset)
                .setIcon(android.R.drawable.ic_menu_close_clear_cancel)
                .setShortcut('2', 'r');
        menu.add(0, MENU_SETTINGS, 0, R.string.settings)
                .setIcon(android.R.drawable.ic_menu_preferences)
                .setShortcut('8', 's')
                .setIntent(new Intent(this, Settings.class));

        if (!mSharedPreferences.getBoolean(PREF_KEY_TWITTER_LOGIN, false)) {
            menu.add(0, MENU_LOGIN, 0, R.string.login)
                    .setShortcut('1', 'l');
        }

        menu.add(0, MENU_QUIT, 0, R.string.quit)
                .setIcon(android.R.drawable.ic_lock_power_off)
                .setShortcut('9', 'q');
        return true;
    }

    /* Handles item selections */
    public boolean onOptionsItemSelected(MenuItem item) {
        switch (item.getItemId()) {
            case MENU_PAUSE:
                pauseGame();
                return true;

            case MENU_RESUME:
                startGame();
                return true;
            case MENU_RESET:
                resetValues(true);
                return true;

            case MENU_CONBT:
                new ConnectBT().execute((Void[]) null);
                return true;

            case MENU_DISCONBT:
                new DisconnectBT().execute((Void[])null);
                return true;

            case MENU_LOGIN:
                loginToTwitter();
                return true;
            case MENU_QUIT:
                resetValues(false);
                unbindStepService();
                stopStepService();
                mQuitting = true;

                if (!(scheduleTaskExecutor == null) && mSharedPreferences.getBoolean(PREF_KEY_TWITTER_LOGIN, false)) {
                    
                    scheduleTaskExecutor.shutdownNow();
                }

                new DisconnectBT().execute();
                if(mBluetoothAdapter.isEnabled())
                    mBluetoothAdapter.disable();

                finish();
                return true;
        }
        return false;
    }

    private void startGame(){
        start_enabled = true;
        startStepService();
        bindStepService();
        mAudio.startRecording(0); // 0 for Hz and 1 for dB

        if (mSharedPreferences.getBoolean(PREF_KEY_TWITTER_LOGIN, false)) {
            scheduleTaskExecutor = Executors.newScheduledThreadPool(5);

            // This schedule a runnable task every 1 minutes
            scheduleTaskExecutor.scheduleAtFixedRate(new Runnable() {
                public void run() {
                    new updateTwitterStatus().execute("Hello there! ");
                    //Log.e("Twitter","Hello there! " + c);
                }
            }, 0, 1, TimeUnit.MINUTES);
        }
    }

    private void pauseGame(){
        start_enabled = false;
        unbindStepService();
        stopStepService();
        mAudio.stopRecording();
        if (scheduleTaskExecutor != null) scheduleTaskExecutor.shutdownNow();
    }

    // TODO: unite all into 1 type of message
    private StepService.ICallback mCallback = new StepService.ICallback() {
        public void stepsChanged(int value) {
            mHandler.sendMessage(mHandler.obtainMessage(STEPS_MSG, value, 0));
        }
        public void paceChanged(int value) {
            mHandler.sendMessage(mHandler.obtainMessage(PACE_MSG, value, 0));
        }
        public void distanceChanged(float value) {
            mHandler.sendMessage(mHandler.obtainMessage(DISTANCE_MSG, (int)(value*1000), 0));
        }
        public void speedChanged(float value) {
            mHandler.sendMessage(mHandler.obtainMessage(SPEED_MSG, (int)(value*1000), 0));
        }
        public void caloriesChanged(float value) {
            mHandler.sendMessage(mHandler.obtainMessage(CALORIES_MSG, (int)(value), 0));
        }
    };

    private static final int STEPS_MSG = 1;
    private static final int PACE_MSG = 2;
    private static final int DISTANCE_MSG = 3;
    private static final int SPEED_MSG = 4;
    private static final int CALORIES_MSG = 5;

    private Handler mHandler = new Handler() {
        @Override public void handleMessage(Message msg) {
            switch (msg.what) {
                case STEPS_MSG:
                    mStepValue = (int)msg.arg1;
                    mStepValueView.setText("" + mStepValue);
                    break;
                case PACE_MSG:
                    mPaceValue = msg.arg1;
                    if (mPaceValue <= 0) {
                        mPaceValueView.setText("0");
                    }
                    else {
                        mPaceValueView.setText("" + (int)mPaceValue);
                    }
                    break;
                case DISTANCE_MSG:
                    mDistanceValue = ((int)msg.arg1)/1000f;
                    if (mDistanceValue <= 0) {
                        mDistanceValueView.setText("0");
                    }
                    else {
                        mDistanceValueView.setText(
                                ("" + (mDistanceValue + 0.000001f)).substring(0, 5)
                        );
                    }
                    break;
                case SPEED_MSG:
                    mSpeedValue = ((int)msg.arg1)/1000f;
                    if (mSpeedValue <= 0) {
                        mSpeedValueView.setText("0");
                    }
                    else {
                        mSpeedValueView.setText(
                                ("" + (mSpeedValue + 0.000001f)).substring(0, 4)
                        );
                    }
                    break;
                case CALORIES_MSG:
                    mCaloriesValue = msg.arg1;
                    if (mCaloriesValue <= 0) {
                        mCaloriesValueView.setText("0");
                    }
                    else {
                        mCaloriesValueView.setText("" + (int)mCaloriesValue);
                    }
                    break;
                default:
                    super.handleMessage(msg);
            }
        }

    };

    /* Reading twitter essential configuration parameters from strings.xml */
    private void initTwitterConfigs() {
        consumerKey = getString(R.string.twitter_consumer_key);
        consumerSecret = getString(R.string.twitter_consumer_secret);
        callbackUrl = getString(R.string.twitter_callback);
        oAuthVerifier = getString(R.string.twitter_oauth_verifier);
    }

    /**
     * Saving user information, after user is authenticated for the first time.
     * You don't need to show user to login, until user has a valid access toen
     */
    private void saveTwitterInfo(AccessToken accessToken) {

        long userID = accessToken.getUserId();

        User user;
        try {
            user = twitter.showUser(userID);

            String username = user.getName();

			/* Storing oAuth tokens to shared preferences */
            SharedPreferences.Editor e = mSharedPreferences.edit();
            e.putString(PREF_KEY_OAUTH_TOKEN, accessToken.getToken());
            e.putString(PREF_KEY_OAUTH_SECRET, accessToken.getTokenSecret());
            e.putBoolean(PREF_KEY_TWITTER_LOGIN, true);
            e.putString(PREF_USER_NAME, username);
            e.commit();

        } catch (TwitterException e1) {
            e1.printStackTrace();
        }
    }

    private void loginToTwitter() {
        boolean isLoggedIn = mSharedPreferences.getBoolean(PREF_KEY_TWITTER_LOGIN, false);

        if (!isLoggedIn) {
            final ConfigurationBuilder builder = new ConfigurationBuilder();
            builder.setOAuthConsumerKey(consumerKey);
            builder.setOAuthConsumerSecret(consumerSecret);

            final Configuration configuration = builder.build();
            final TwitterFactory factory = new TwitterFactory(configuration);
            twitter = factory.getInstance();

            try {
                requestToken = twitter.getOAuthRequestToken(callbackUrl);

                /**
                 *  Loading twitter login page on webview for authorization
                 *  Once authorized, results are received at onActivityResult
                 *  */
                final Intent intent = new Intent(this, WebViewActivity.class);
                intent.putExtra(WebViewActivity.EXTRA_URL, requestToken.getAuthenticationURL());
                startActivityForResult(intent, WEBVIEW_REQUEST_CODE);

            } catch (TwitterException e) {
                e.printStackTrace();
            }
        } else {

            /*loginLayout.setVisibility(View.GONE);
            shareLayout.setVisibility(View.VISIBLE);*/
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {

        if (resultCode == Activity.RESULT_OK) {
            String verifier = data.getExtras().getString(oAuthVerifier);
            try {
                AccessToken accessToken = twitter.getOAuthAccessToken(requestToken, verifier);

                long userID = accessToken.getUserId();
                final User user = twitter.showUser(userID);
                String username = user.getName();

                saveTwitterInfo(accessToken);

                /*loginLayout.setVisibility(View.GONE);
                shareLayout.setVisibility(View.VISIBLE);
                userName.setText(MainActivity.this.getResources().getString(
                        R.string.hello) + username);*/

            } catch (Exception e) {
                Log.e("Twitter Login Failed", e.getMessage());
            }
        }

        super.onActivityResult(requestCode, resultCode, data);
    }

    class updateTwitterStatus extends AsyncTask<String, String, Void> {
        @Override
        protected void onPreExecute() {
            super.onPreExecute();

			/*pDialog = new ProgressDialog(MainActivity.this);
			pDialog.setMessage("Posting to twitter...");
			pDialog.setIndeterminate(false);
			pDialog.setCancelable(false);
			pDialog.show();*/
        }

        protected Void doInBackground(String... args) {

            String status = args[0];
            try {
                ConfigurationBuilder builder = new ConfigurationBuilder();
                builder.setOAuthConsumerKey(consumerKey);
                builder.setOAuthConsumerSecret(consumerSecret);

                // Access Token
                String access_token = mSharedPreferences.getString(PREF_KEY_OAUTH_TOKEN, "");
                // Access Token Secret
                String access_token_secret = mSharedPreferences.getString(PREF_KEY_OAUTH_SECRET, "");

                AccessToken accessToken = new AccessToken(access_token, access_token_secret);
                Twitter twitter = new TwitterFactory(builder.build()).getInstance(accessToken);

                // Update status yyyy/MM/dd HH:mm:ss
                int respirationRate=0;
                SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy-HH:mm:ss");
                if (mStepValue != 0) {
                    respirationRate = mStepValue/2;
                }
                String date = sdf.format(new Date());
                StringBuilder sb = new StringBuilder();
                sb.append("RFID: ").append(rfidList.get(wayPoint-1)).append(" Time: ").append(date)
                    .append(" HeartRate: ").append(heartRateGlobal).append(" [bpm]")
                    .append(" RespirationRate: ").append(respirationRate)
                    .append(" Location: ").append(myLocation.getLatitude()).append(" lat ").append(myLocation.getLongitude())
                    .append(" long");
                StatusUpdate statusUpdate = new StatusUpdate(sb.toString());

                // add image if you want
				/*InputStream is = getResources().openRawResource(R.drawable.lakeside_view);
				statusUpdate.setMedia("test.jpg", is);*/

                twitter4j.Status response = twitter.updateStatus(statusUpdate);

                Log.d("Status", response.getText());

            } catch (TwitterException e) {
                Log.d("Failed to post!", e.getMessage());
            }
            return null;
        }

        @Override
        protected void onPostExecute(Void result) {

			/* Dismiss the progress dialog after sharing */
            //pDialog.dismiss();

            Toast.makeText(MainActivity.this, "Posted to Twitter!", Toast.LENGTH_SHORT).show();

            // Clearing EditText field
            // mShareEditText.setText("");

        }

    }

    private class ConnectBT extends AsyncTask<Void,Void,Void > {
        ProgressDialog dialog;
        @Override
        protected void onPreExecute() {
            dialog = new ProgressDialog(MainActivity.this);
            dialog.setMessage("Connecting to device...");
            dialog.setIndeterminate(true);
            dialog.setCancelable(false);
            dialog.show();
        }

        @Override
        protected Void doInBackground(Void... params) {

            try {
                connectBTDevice();
            } catch (IOException e1) {
                e1.printStackTrace();
            }

            return null;
        }

        protected void onPostExecute(Void result) {
            dialog.dismiss();
            if (mmSocket.isConnected()) {
                listenAndEvolve();
            }
        }
    }

    private class DisconnectBT extends AsyncTask<Void,Void,Void > {
        ProgressDialog dialog;
        @Override
        protected void onPreExecute() {
            dialog = new ProgressDialog(MainActivity.this);
            dialog.setMessage("Disconnecting from device...");
            dialog.setIndeterminate(true);
            dialog.setCancelable(false);
            dialog.show();
        }

        @Override
        protected Void doInBackground(Void... params) {

            try {
                disconnectBTDevice();
            } catch (IOException ex) {
                ex.printStackTrace();
            }

            return null;
        }

        protected void onPostExecute(Void result) {
            dialog.dismiss();
        }
    }

    private boolean saveExcelStart(Context context, String fileName) {

        // check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e(TAG, "Storage not available or read only");
            return false;
        }

        boolean success = false;

        //New Workbook
        Workbook wb = new HSSFWorkbook();

        Cell c = null;

        //Cell style for header row
        Font headerFont = wb.createFont();
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cs.setFont(headerFont);

        CellStyle cs2 = wb.createCellStyle();
        cs2.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        cs2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //New Sheet
        Sheet sheet1 = null;
        sheet1 = wb.createSheet("H&S2015");

        // Generate column headings
        Row row = sheet1.createRow(0);

        c = row.createCell(0);
        c.setCellValue("Team ID");
        c.setCellStyle(cs);

        c = row.createCell(1);
        c.setCellValue("Suceava2");
        //c.setCellStyle(cs);

        c = row.createCell(2);
        c.setCellValue("");
        c.setCellStyle(cs);


        // Generate column headings
        Row row2 = sheet1.createRow(1);

        c = row2.createCell(0);
        c.setCellValue("");
        c.setCellStyle(cs);

        c = row2.createCell(1);
        c.setCellValue("Time Stamp (dd.mm.yy - hh.mm.ss)");
        c.setCellStyle(cs);

        c = row2.createCell(2);
        c.setCellValue("Recorded values");
        c.setCellStyle(cs);

        // -------------------------------
        Row row3 = sheet1.createRow(2);

        c = row3.createCell(0);
        c.setCellValue("Start information");
        c.setCellStyle(cs);

        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy-HH:mm:ss");
        String date = sdf.format(new Date());
        c = row3.createCell(1);
        c.setCellValue(date);
        //c.setCellStyle(cs);

        c = row3.createCell(2);
        c.setCellValue("");
        c.setCellStyle(cs);

        // -------------------------------
        Row row4 = sheet1.createRow(3);

        c = row4.createCell(0);
        c.setCellValue("    Bio-data");
        c.setCellStyle(cs2);

        c = row4.createCell(1);
        c.setCellValue(date);
        //c.setCellStyle(cs2);

        c = row4.createCell(2);
        c.setCellValue("");
        c.setCellStyle(cs2);

        // -------------------------------
        Row row5 = sheet1.createRow(4);

        c = row5.createCell(0);
        c.setCellValue("        Body temperature [°C]");
        c.setCellStyle(cs2);

        c = row5.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs2);

        c = row5.createCell(2);
        c.setCellValue(String.valueOf(bodyTempGlobal));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row6 = sheet1.createRow(5);

        c = row6.createCell(0);
        c.setCellValue("        Pulse rate [bpm]");
        c.setCellStyle(cs2);

        c = row6.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs2);

        c = row6.createCell(2);
        c.setCellValue(String.valueOf(heartRateGlobal));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row7 = sheet1.createRow(6);

        c = row7.createCell(0);
        c.setCellValue("        Number of steps [decimal]");
        c.setCellStyle(cs2);

        c = row7.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs2);

        c = row7.createCell(2);
        c.setCellValue(String.valueOf(mStepValue));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row8 = sheet1.createRow(7);

        c = row8.createCell(0);
        c.setCellValue("        Distance traveled [m]");
        c.setCellStyle(cs2);

        c = row8.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs2);

        c = row8.createCell(2);
        c.setCellValue(String.valueOf(mDistanceValue * 1000));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row9 = sheet1.createRow(8);

        c = row9.createCell(0);
        c.setCellValue("        Burned calories");
        c.setCellStyle(cs2);

        c = row9.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs2);

        c = row9.createCell(2);
        c.setCellValue(String.valueOf((int)mCaloriesValue));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row10 = sheet1.createRow(9);

        c = row10.createCell(0);
        c.setCellValue("    Environmental data");
        c.setCellStyle(cs2);

        c = row10.createCell(1);
        c.setCellValue(date);
        //c.setCellStyle(cs);

        c = row10.createCell(2);
        c.setCellValue("");
        c.setCellStyle(cs2);

        // -------------------------------
        Row row11= sheet1.createRow(10);

        c = row11.createCell(0);
        c.setCellValue("        Air temperature [°C]");
        c.setCellStyle(cs2);

        c = row11.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs);

        c = row11.createCell(2);
        c.setCellValue(String.valueOf(airTempGlobal));
        //c.setCellStyle(cs2);

        // -------------------------------
        Row row12 = sheet1.createRow(11);

        c = row12.createCell(0);
        c.setCellValue("        Humidity [%]");
        c.setCellStyle(cs2);

        c = row12.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs);

        c = row12.createCell(2);
        c.setCellValue(String.valueOf(humidityGlobal));
        //c.setCellStyle(cs2);

        // ------------------------
        Row row13 = sheet1.createRow(12);

        c = row13.createCell(0);
        c.setCellValue("        Ambient light [%]");
        c.setCellStyle(cs2);

        c = row13.createCell(1);
        c.setCellValue("");
        c.setCellStyle(cs);

        c = row13.createCell(2);
        c.setCellValue(String.valueOf(ambiLightGlobal));
        //c.setCellStyle(cs2);


        // -------------------/////-------------------
        sheet1.setColumnWidth(0, (15 * 530));
        sheet1.setColumnWidth(1, (15 * 590));
        sheet1.setColumnWidth(2, (15 * 500));

        // Create a path where we will place our List of objects on external storage
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return success;
    }

    private boolean readWriteStation(Context context, int idStation, String filename) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.e(TAG, "Storage not available or read only");
            return false;
        }
        boolean success = false;

        try{
            // Creating Input Stream
            File file = new File(context.getExternalFilesDir(null), filename);
            FileInputStream myInput = new FileInputStream(file);

            // Create a POIFSFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            // Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            // Get the first sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);

            /** We now need something to iterate through the cells.**/
            Iterator rowIter = mySheet.rowIterator();
            int cnt=0;

            while(rowIter.hasNext()){
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator cellIter = myRow.cellIterator();
                /*while(cellIter.hasNext()){
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    Log.d(TAG, "Cell Value: " +  myCell.toString());
                    //Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                }*/
                //Log.e("ExelLog","Row: " + cnt++);
                cnt++;
            }

            // lets write something

            //Cell style for header row
            Font headerFont = myWorkBook.createFont();
            headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cs = myWorkBook.createCellStyle();
            cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cs.setFont(headerFont);

            CellStyle cs2 = myWorkBook.createCellStyle();
            cs2.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            cs2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

            // Generate column headings
            Row row1 = mySheet.createRow(cnt++);
            Cell c;

            c = row1.createCell(0);
            c.setCellValue("Station " + idStation);
            c.setCellStyle(cs);

            SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy-HH:mm:ss");
            String date = sdf.format(new Date());
            c = row1.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs);

            c = row1.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs);

            // ---------------------------
            Row row2 = mySheet.createRow(cnt++);

            c = row2.createCell(0);
            c.setCellValue("    Bio-data");
            c.setCellStyle(cs2);

            c = row2.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs);

            c = row2.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // ---------------------------
            Row row3 = mySheet.createRow(cnt++);

            c = row3.createCell(0);
            c.setCellValue("        Body temperature [°C]");
            c.setCellStyle(cs2);

            c = row3.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row3.createCell(2);
            c.setCellValue(String.valueOf(bodyTempGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row4 = mySheet.createRow(cnt++);

            c = row4.createCell(0);
            c.setCellValue("        Pulse rate [bpm]");
            c.setCellStyle(cs2);

            c = row4.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row4.createCell(2);
            c.setCellValue(String.valueOf(heartRateGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row5 = mySheet.createRow(cnt++);

            c = row5.createCell(0);
            c.setCellValue("        Number of steps [decimal]");
            c.setCellStyle(cs2);

            c = row5.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row5.createCell(2);
            c.setCellValue(String.valueOf(mStepValue));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row6 = mySheet.createRow(cnt++);

            c = row6.createCell(0);
            c.setCellValue("        Distance traveled [m]");
            c.setCellStyle(cs2);

            c = row6.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row6.createCell(2);
            c.setCellValue(String.valueOf(mDistanceValue * 1000));
            //c.setCellStyle(cs2);

            // -------------------------------
            Row rowX = mySheet.createRow(cnt++);

            c = rowX.createCell(0);
            c.setCellValue("        Burned calories");
            c.setCellStyle(cs2);

            c = rowX.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = rowX.createCell(2);
            c.setCellValue(String.valueOf((int)mCaloriesValue));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row7 = mySheet.createRow(cnt++);

            c = row7.createCell(0);
            c.setCellValue("    Environmental data");
            c.setCellStyle(cs2);

            c = row7.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs2);

            c = row7.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // ---------------------------
            Row row8 = mySheet.createRow(cnt++);

            c = row8.createCell(0);
            c.setCellValue("        Air temperature [°C]");
            c.setCellStyle(cs2);

            c = row8.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row8.createCell(2);
            c.setCellValue(String.valueOf(airTempGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row9 = mySheet.createRow(cnt++);

            c = row9.createCell(0);
            c.setCellValue("        Humidity [%]");
            c.setCellStyle(cs2);

            c = row9.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row9.createCell(2);
            c.setCellValue(String.valueOf(humidityGlobal));
            //c.setCellStyle(cs2);

            // ------------------------
            Row row13 = mySheet.createRow(cnt++);

            c = row13.createCell(0);
            c.setCellValue("        Ambient light [%]");
            c.setCellStyle(cs2);

            c = row13.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs);

            c = row13.createCell(2);
            c.setCellValue(String.valueOf(ambiLightGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row10 = mySheet.createRow(cnt++);

            c = row10.createCell(0);
            c.setCellValue("Treasure records");
            c.setCellStyle(cs2);

            c = row10.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs2);

            c = row10.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // --------------------------- TODO de revazut

            String a[][] = {{"0","0"},{"0","0"},{"0","0"},{"0","0"},{"0","0"}};
            if (remoteTreasureGlobal.length() > 1)  { a[0][0]="1"; a[0][1]="Remote: " + remoteTreasureGlobal; }
            if (airTempTreasureGlobal > 35)         { a[1][0]="1"; a[1][1]="Temp: " + String.valueOf(airTempTreasureGlobal); }
            if (magneticTreasureGlobal == 1)        { a[2][0]="1"; a[2][1]="Magnetic field found"; }
            if (blinkRateGlobal != 0)               { a[3][0]="1"; a[3][1]="BlinkRate: " + String.valueOf(blinkRateGlobal); }
            if (frequencyGlobal != 0)               { a[4][0]="1"; a[4][1]="Frequency: " + String.valueOf((int)frequencyGlobal); }

            int counter = 0;
            for (int i = 0; i < 5; i++) {
                if (a[i][0].equals("1")) {
                    ++counter;
                }
            }


            Row row11 = mySheet.createRow(cnt++);

            c = row11.createCell(0);
            c.setCellValue("        Number of treasures");
            c.setCellStyle(cs2);

            c = row11.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs2);

            c = row11.createCell(2);
            c.setCellValue(counter);
            //c.setCellStyle(cs2);

            // --------------------------- TODO de revazut
            int knt=0;
            for (int i = 0; i < 5; i++) {
                if (i==0) knt = 0;
                if (a[i][0].equals("1")) {
                    Row row12 = mySheet.createRow(cnt++);

                    c = row12.createCell(0);
                    c.setCellValue("        Treasure " + (++knt));
                    c.setCellStyle(cs2);

                    c = row12.createCell(1);
                    c.setCellValue(date);
                    //c.setCellStyle(cs2);

                    c = row12.createCell(2);
                    c.setCellValue(a[i][1]);
                    //c.setCellStyle(cs2);
                }

            }

            // -----------------////--------------
            mySheet.setColumnWidth(0, (15 * 530));
            mySheet.setColumnWidth(1, (15 * 590));
            mySheet.setColumnWidth(2, (15 * 500));

            // Create a path where we will place our List of objects on external storage
            File file2 = new File(context.getExternalFilesDir(null), filename);
            FileOutputStream os = null;

            try {
                os = new FileOutputStream(file2);
                myWorkBook.write(os);
                Log.w("FileUtils", "Writing file" + file2);
                success = true;
            } catch (IOException e) {
                Log.w("FileUtils", "Error writing " + file2, e);
            } catch (Exception e) {
                Log.w("FileUtils", "Failed to save file", e);
            } finally {
                try {
                    if (null != os)
                        os.close();
                } catch (Exception ex) {
                }
            }


        }catch (Exception e){e.printStackTrace(); }

        return success;
    }

    private boolean saveExcelFinish(Context context, String filename) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.e(TAG, "Storage not available or read only");
            return false;
        }
        boolean success = false;

        try{
            // Creating Input Stream
            File file = new File(context.getExternalFilesDir(null), filename);
            FileInputStream myInput = new FileInputStream(file);

            // Create a POIFSFileSystem object
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            // Create a workbook using the File System
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            // Get the first sheet from workbook
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);

            /** We now need something to iterate through the cells.**/
            Iterator rowIter = mySheet.rowIterator();
            int cnt=0;

            while(rowIter.hasNext()){
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator cellIter = myRow.cellIterator();
                /*while(cellIter.hasNext()){
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    Log.d(TAG, "Cell Value: " +  myCell.toString());
                    //Toast.makeText(context, "cell Value: " + myCell.toString(), Toast.LENGTH_SHORT).show();
                }*/
                //Log.e("ExelLog","Row: " + cnt++);
                cnt++;
            }

            // lets write something

            //Cell style for header row
            Font headerFont = myWorkBook.createFont();
            headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            CellStyle cs = myWorkBook.createCellStyle();
            cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cs.setFont(headerFont);

            CellStyle cs2 = myWorkBook.createCellStyle();
            cs2.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            cs2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

            // Generate column headings
            Row row1 = mySheet.createRow(cnt++);
            Cell c;

            c = row1.createCell(0);
            c.setCellValue("Finish line");
            c.setCellStyle(cs);

            SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy-HH:mm:ss");
            String date = sdf.format(new Date());
            c = row1.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs);

            c = row1.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs);

            // ---------------------------
            Row row2 = mySheet.createRow(cnt++);

            c = row2.createCell(0);
            c.setCellValue("    Bio-data");
            c.setCellStyle(cs2);

            c = row2.createCell(1);
            c.setCellValue("date");
            //c.setCellStyle(cs);

            c = row2.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // ---------------------------
            Row row3 = mySheet.createRow(cnt++);

            c = row3.createCell(0);
            c.setCellValue("        Body temperature [°C]");
            c.setCellStyle(cs2);

            c = row3.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row3.createCell(2);
            c.setCellValue(String.valueOf(bodyTempGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row4 = mySheet.createRow(cnt++);

            c = row4.createCell(0);
            c.setCellValue("        Pulse rate [bpm]");
            c.setCellStyle(cs2);

            c = row4.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row4.createCell(2);
            c.setCellValue(String.valueOf(heartRateGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row5 = mySheet.createRow(cnt++);

            c = row5.createCell(0);
            c.setCellValue("        Number of steps [decimal]");
            c.setCellStyle(cs2);

            c = row5.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row5.createCell(2);
            c.setCellValue(String.valueOf(mStepValue));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row6 = mySheet.createRow(cnt++);

            c = row6.createCell(0);
            c.setCellValue("        Distance traveled [m]");
            c.setCellStyle(cs2);

            c = row6.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row6.createCell(2);
            c.setCellValue(String.valueOf(mDistanceValue * 1000));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row7 = mySheet.createRow(cnt++);

            c = row7.createCell(0);
            c.setCellValue("    Environmental data");
            c.setCellStyle(cs2);

            c = row7.createCell(1);
            c.setCellValue(date);
            //c.setCellStyle(cs2);

            c = row7.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // ---------------------------
            Row row8 = mySheet.createRow(cnt++);

            c = row8.createCell(0);
            c.setCellValue("        Air temperature [°C]");
            c.setCellStyle(cs2);

            c = row8.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row8.createCell(2);
            c.setCellValue(String.valueOf(airTempGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row9 = mySheet.createRow(cnt++);

            c = row9.createCell(0);
            c.setCellValue("        Humidity [%]");
            c.setCellStyle(cs2);

            c = row9.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row9.createCell(2);
            c.setCellValue(String.valueOf(humidityGlobal));
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row10 = mySheet.createRow(cnt++);

            c = row10.createCell(0);
            c.setCellValue("Total identified treasures");
            c.setCellStyle(cs2);

            c = row10.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row10.createCell(2);
            c.setCellValue("");
            c.setCellStyle(cs2);

            // ---------------------------
            Row row11 = mySheet.createRow(cnt++);

            c = row11.createCell(0);
            c.setCellValue("Total time (start - finish)");
            c.setCellStyle(cs2);

            c = row11.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row11.createCell(2);
            c.setCellValue("value");
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row12 = mySheet.createRow(cnt++);

            c = row12.createCell(0);
            c.setCellValue("Min. Pulse rate");
            c.setCellStyle(cs2);

            c = row12.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row12.createCell(2);
            c.setCellValue("value");
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row13 = mySheet.createRow(cnt++);

            c = row13.createCell(0);
            c.setCellValue("Max. Pulse rate");
            c.setCellStyle(cs2);

            c = row13.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row13.createCell(2);
            c.setCellValue("value");
            //c.setCellStyle(cs2);

            // ---------------------------
            Row row14 = mySheet.createRow(cnt++);

            c = row14.createCell(0);
            c.setCellValue("Average pulse rate");
            c.setCellStyle(cs2);

            c = row14.createCell(1);
            c.setCellValue("");
            c.setCellStyle(cs2);

            c = row14.createCell(2);
            c.setCellValue("value");
            //c.setCellStyle(cs2);


            // -----------------////--------------
            mySheet.setColumnWidth(0, (15 * 530));
            mySheet.setColumnWidth(1, (15 * 590));
            mySheet.setColumnWidth(2, (15 * 500));

            // Create a path where we will place our List of objects on external storage
            File file2 = new File(context.getExternalFilesDir(null), filename);
            FileOutputStream os = null;

            try {
                os = new FileOutputStream(file2);
                myWorkBook.write(os);
                Log.w("FileUtils", "Writing file" + file2);
                success = true;
            } catch (IOException e) {
                Log.w("FileUtils", "Error writing " + file2, e);
            } catch (Exception e) {
                Log.w("FileUtils", "Failed to save file", e);
            } finally {
                try {
                    if (null != os)
                        os.close();
                } catch (Exception ex) {
                }
            }


        }catch (Exception e) {e.printStackTrace(); }

        return success;
    }


    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }

}