import { useEffect, useRef, useState } from "react";

interface CaptureControlsProps {
  onRecordingFinished: (blob: Blob) => void;
}

const CaptureControls = ({ onRecordingFinished }: CaptureControlsProps) => {
  const [isRecording, setIsRecording] = useState(false);
  const [isSupported, setIsSupported] = useState(true);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const chunksRef = useRef<Blob[]>([]);
  const videoRef = useRef<HTMLVideoElement | null>(null);

  useEffect(() => {
    if (!navigator.mediaDevices?.getUserMedia) {
      setIsSupported(false);
    }
  }, []);

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: true,
        video: true
      });
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
      }
      const mediaRecorder = new MediaRecorder(stream);
      mediaRecorderRef.current = mediaRecorder;
      chunksRef.current = [];
      mediaRecorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          chunksRef.current.push(event.data);
        }
      };
      mediaRecorder.onstop = () => {
        const blob = new Blob(chunksRef.current, { type: "video/webm" });
        onRecordingFinished(blob);
        stream.getTracks().forEach((track) => track.stop());
      };
      mediaRecorder.start();
      setIsRecording(true);
    } catch (error) {
      console.error("No se pudo iniciar la grabación", error);
      setIsSupported(false);
    }
  };

  const stopRecording = () => {
    mediaRecorderRef.current?.stop();
    setIsRecording(false);
  };

  return (
    <section className="card">
      <header>
        <h2>Captura de audio y vídeo</h2>
        <p>Controla el inicio y fin de las grabaciones locales.</p>
      </header>
      {!isSupported ? (
        <p className="warning">
          Tu navegador no soporta captura multimedia. Utiliza Chrome o Edge reciente.
        </p>
      ) : (
        <>
          <div className="capture-actions">
            <button onClick={startRecording} disabled={isRecording} type="button">
              Iniciar captura
            </button>
            <button onClick={stopRecording} disabled={!isRecording} type="button">
              Detener
            </button>
          </div>
          <video ref={videoRef} muted playsInline className="preview" />
        </>
      )}
    </section>
  );
};

export default CaptureControls;
