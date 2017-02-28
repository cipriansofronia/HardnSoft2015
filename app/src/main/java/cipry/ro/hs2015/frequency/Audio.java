package cipry.ro.hs2015.frequency;

import android.media.AudioFormat;
import android.media.AudioRecord;
import android.media.MediaRecorder;
import android.os.Bundle;
import android.os.Handler;
import android.os.Message;
import android.util.Log;

import edu.emory.mathcs.jtransforms.fft.DoubleFFT_1D;

public class Audio implements Runnable  {

    private AudioRecord mAudioRecord = null;
    private Thread mThread = null;
    private Handler mHandler = null;
    private int BUFFER_SIZE = 1;
    private boolean RECORD;
    private int DATA_TYPE = 0;

    public Audio(Handler handler){
        try{
            mHandler = handler;
            BUFFER_SIZE = 4*mAudioRecord.getMinBufferSize(8000 /*MediaRecorder.AudioSource.MIC*/, AudioFormat.CHANNEL_IN_MONO, AudioFormat.ENCODING_PCM_16BIT);
            mAudioRecord = new AudioRecord(MediaRecorder.AudioSource.MIC, 44100, AudioFormat.CHANNEL_IN_MONO, AudioFormat.ENCODING_PCM_16BIT, BUFFER_SIZE);
            System.out.println("BUFFER SIZE VALUE IS " + BUFFER_SIZE);
        }
        catch (Exception e){
            Log.d("ERROR",e.toString());
        }
    }

    @Override
    public void run() {
        DoubleFFT_1D FFT = new DoubleFFT_1D(BUFFER_SIZE);
        short [] realAudioData = new short [BUFFER_SIZE];
        double [] doubleAudioData;
        double [] magnitude;
        double frequency = 0;
        double dB = 0;
        mAudioRecord.startRecording();
        
        android.os.Process.setThreadPriority(android.os.Process.THREAD_PRIORITY_BACKGROUND);
        while (RECORD){
            
            mAudioRecord.read(realAudioData,0,BUFFER_SIZE);
            
            doubleAudioData = convertShortToDoubleArray(realAudioData);
            
            FFT.realForward(doubleAudioData);
            
            magnitude = magnitude(doubleAudioData);
            if (DATA_TYPE == 0){
                 
                int peakIndeks = peakIndeks(magnitude);
                frequency = calculateFrequecny(peakIndeks);
            }
            else{
                
                dB = calculateDecibels(avrageSignalPower(realAudioData));
            }
            
            sendMessageToUI(frequency,dB);
        }
        mAudioRecord.stop();
    }

    public void sendMessageToUI(double frequency,double dB){
        Message message = new Message();
        Bundle bundle = new Bundle();
        bundle.putDouble("frequency",frequency);
        bundle.putDouble("dB",dB);
        message.setData(bundle);
        mHandler.sendMessage(message);
    }

    public void startRecording(int type){
        DATA_TYPE = type;
        RECORD = true;
        mThread = new Thread(this);
        mThread.start();
    }

    public void stopRecording(){
        RECORD = false;
    }

    private double [] convertShortToDoubleArray(short [] originalData){
        double [] newData = new double[2*originalData.length];
        for (int i=0;i<originalData.length;i++){
            newData[i] = (double)originalData[i];
        }
        return  newData;
    }

    private double [] magnitude(double [] FFTdata){
        double [] magnitude = new double[FFTdata.length/2];
        for (int i=0;i<FFTdata.length/2;i++){
            magnitude[i] = Math.sqrt(FFTdata[2*i]*FFTdata[2*i]+FFTdata[2*i+1]*FFTdata[2*i+1]);
        }
        return  magnitude;
    }

    private int peakIndeks(double [] FFTdata){
        int index = -1;
        double maxMagnitude = Double.NEGATIVE_INFINITY;
        for (int i=0;i<FFTdata.length/2;i++){
            if (maxMagnitude<Math.abs(FFTdata[i])){
                index = i;
                maxMagnitude = Math.abs(FFTdata[i]);
            }
        }
        return index;
    }

    private double calculateFrequecny(int index){
        double frequency = index * (44100.0 / BUFFER_SIZE);
        return frequency;
    }

    private double calculateDecibels(double avgSignalPower){
        double dB = 20.0 * Math.log10(avgSignalPower/Short.MAX_VALUE);
        return  dB;
    }

    private double avrageSignalPower(short [] amplitudes){
        double avgSignalPower = 0;
        float avgAmplitudes = 0;
        for (int i=0;i<amplitudes.length;i++){
            avgAmplitudes += amplitudes[i]*amplitudes[i];
        }
        avgSignalPower = (double)(avgAmplitudes/amplitudes.length);
        return Math.sqrt(avgSignalPower);
    }

    public void release(){
        mAudioRecord.release();
    }
}
