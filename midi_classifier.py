import pandas as pd
import mido
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, timedelta
import json
import os
from typing import Dict, List
import threading
import time
import sys
from queue import Queue

class MIDIPlayer:
    def __init__(self):
        self.current_midi = None
        self.is_playing = False
        self.is_looping = False
        self.loop_start = 0
        self.loop_end = None
        self.tempo = 1.0
        self.current_position = 0
        self.play_thread = None
        self.play_queue = Queue()
        
        # Try to get default port
        try:
            self.port_name = mido.get_output_names()[0]
            self.port = mido.open_output(self.port_name)
            print(f"Using MIDI port: {self.port_name}")
        except IndexError:
            messagebox.showerror("Error", "No MIDI output ports found. Please connect a MIDI device.")
            sys.exit(1)
        
    def load_midi(self, filepath: str):
        """Load a MIDI file"""
        self.stop()
        try:
            self.current_midi = mido.MidiFile(filepath)
            print(f"Loaded MIDI file: {filepath}")
            # Calculate total time
            self.loop_end = sum(msg.time for track in self.current_midi.tracks 
                              for msg in track)
        except Exception as e:
            raise Exception(f"Failed to load MIDI file: {str(e)}")
        
    def play(self, start_pos=None):
        """Start playing the MIDI file"""
        if not self.current_midi:
            return
            
        if self.is_playing:
            self.stop()
            
        self.is_playing = True
        if start_pos is not None:
            self.current_position = start_pos
            
        self.play_thread = threading.Thread(target=self._play_thread)
        self.play_thread.daemon = True  # Thread will close with main program
        self.play_thread.start()
        
    def _play_thread(self):
        """Thread handler for MIDI playback"""
        try:
            while self.is_playing:
                for msg in self.current_midi.play(meta_messages=True):
                    if not self.is_playing:
                        break
                        
                    if hasattr(msg, 'time'):
                        time.sleep(msg.time * (1/self.tempo))
                        
                    if not hasattr(msg, 'type'):
                        continue
                        
                    # Handle only note messages
                    if msg.type in ['note_on', 'note_off']:
                        self.port.send(msg)
                        
                    self.current_position += msg.time
                    if self.is_looping and self.current_position >= self.loop_end:
                        self.current_position = self.loop_start
                        break
                        
                if not self.is_looping:
                    break
                    
        except Exception as e:
            print(f"Playback error: {str(e)}")
            self.is_playing = False
        finally:
            # Send all notes off
            self.stop()
            
    def stop(self):
        """Stop MIDI playback and clear all notes"""
        self.is_playing = False
        if hasattr(self, 'port') and self.port:
            # Send all notes off
            for channel in range(16):
                for note in range(128):
                    self.port.send(mido.Message('note_off', note=note, channel=channel))
        
    def set_tempo(self, tempo: float):
        """Set playback tempo"""
        self.tempo = max(0.25, min(2.0, tempo))
        
    def set_loop_points(self, start: float, end: float):
        """Set loop points for playback"""
        self.loop_start = start
        self.loop_end = end
        
    def toggle_loop(self):
        """Toggle looping playback"""
        self.is_looping = not self.is_looping
        
    def __del__(self):
        """Cleanup when object is destroyed"""
        if hasattr(self, 'port') and self.port:
            self.stop()
            self.port.close()


class ClassificationStats:
    def __init__(self):
        self.stats = {
            'OK': 0,
            'NG1': 0, 'NG2': 0, 'NG3': 0, 'NG4': 0,
            'NG5': 0, 'NG6': 0, 'NG7': 0, 'NG8': 0
        }
        self.total_time = timedelta()
        
    def update(self, classification: str, time_spent: timedelta):
        self.stats[classification] += 1
        self.total_time += time_spent
        
    def get_summary(self) -> Dict:
        total = sum(self.stats.values())
        return {
            'total_files': total,
            'ok_ratio': self.stats['OK'] / total if total else 0,
            'most_common_ng': max(
                [(k, v) for k, v in self.stats.items() if k.startswith('NG')],
                key=lambda x: x[1]
            )[0] if total else None,
            'avg_time_per_file': self.total_time / total if total else timedelta()
        }

class MIDIClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MIDI Melody Classifier")
        
        # Initialize components
        self.player = MIDIPlayer()
        self.stats = ClassificationStats()
        self.classifications = []
        self.current_index = 0
        self.start_time = None
        self.midi_files = []  # List to store MIDI file paths
        
        # Create directory and load files before UI setup
        self.initialize_directory()
        self.load_midi_files()  # Load MIDI files from directory
        self.setup_ui()
        self.setup_shortcuts()
        self.load_progress()

        # Load first file if available
        if self.midi_files:
            self.load_file(self.midi_files[0])

    def initialize_directory(self):
        """Create midi_files directory if it doesn't exist"""
        midi_dir = 'midi_files'
        if not os.path.exists(midi_dir):
            os.makedirs(midi_dir)
            messagebox.showinfo("Directory Created", 
                              "The 'midi_files' directory has been created at:\n"
                              f"{os.path.abspath(midi_dir)}\n"
                              "Please add your MIDI files there and restart the application.")

    def load_midi_files(self):
        """Load all MIDI files from the midi_files directory"""
        midi_dir = 'midi_files'
        if not os.path.exists(midi_dir):
            os.makedirs(midi_dir)
            messagebox.showinfo("Directory Created", 
                              "The 'midi_files' directory has been created. Please add your MIDI files there.")
            return
            
        self.midi_files = [
            os.path.join(midi_dir, f) for f in os.listdir(midi_dir)
            if f.lower().endswith(('.mid', '.midi'))
        ]
        self.midi_files.sort()
        
        if not self.midi_files:
            messagebox.showinfo("No Files Found", 
                              "No MIDI files found in the 'midi_files' directory.")
            
    def prev_file(self):
        """Load the previous MIDI file"""
        if not self.midi_files:
            return
            
        self.player.stop()
        self.current_index = (self.current_index - 1) % len(self.midi_files)
        self.load_file(self.midi_files[self.current_index])
        
    def next_file(self):
        """Load the next MIDI file"""
        if not self.midi_files:
            return
            
        self.player.stop()
        self.current_index = (self.current_index + 1) % len(self.midi_files)
        self.load_file(self.midi_files[self.current_index])
        
    def load_file(self, filepath: str):
        """Load a MIDI file and update the UI"""
        if not filepath or not os.path.exists(filepath):
            self.status_label.config(
                text=f"Error: File not found - {filepath}",
                foreground="red"
            )
            return
            
        try:
            self.player.load_midi(filepath)
            self.current_midi = filepath
            
            # Update labels
            filename = os.path.basename(filepath)
            total_files = len(self.midi_files)
            current_num = self.current_index + 1
            
            self.file_label.config(
                text=f"File {current_num}/{total_files}: {filename}"
            )
            self.status_label.config(
                text=f"Currently processing: {filename}",
                foreground="green"
            )
            
            self.comments_text.delete('1.0', 'end')
            self.start_time = datetime.now()
            
        except Exception as e:
            self.status_label.config(
                text=f"Error loading file: {str(e)}",
                foreground="red"
            )
            messagebox.showerror("Error", f"Failed to load MIDI file: {e}")
            
    def refresh_files(self):
        """Refresh the list of MIDI files"""
        self.load_midi_files()
        if self.midi_files:
            self.current_index = 0
            self.load_file(self.midi_files[0])
        
    def setup_ui(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(expand=True, fill="both")
        
        # Status frame at the top
        status_frame = ttk.LabelFrame(main_container, text="Status")
        status_frame.pack(fill="x", pady=(0, 10))
        
        self.status_label = ttk.Label(status_frame, text="Loading...")
        self.status_label.pack(pady=5)
        
        # Update status based on file loading
        if not self.midi_files:
            self.status_label.config(
                text="No MIDI files found. Please add files to the 'midi_files' directory.",
                foreground="red"
            )
        else:
            self.status_label.config(
                text=f"Loaded {len(self.midi_files)} MIDI files",
                foreground="green"
            )
        
        # Left panel (File info and controls)
        left_panel = ttk.Frame(main_container)
        left_panel.pack(side="left", fill="both", expand=True)
        
        # File information
        file_frame = ttk.LabelFrame(left_panel, text="Current File")
        file_frame.pack(fill="x", pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No file loaded")
        self.file_label.pack(pady=5)
        
        # Playback controls
        playback_frame = ttk.LabelFrame(left_panel, text="Playback Controls")
        playback_frame.pack(fill="x", pady=5)
        
        ttk.Button(playback_frame, text="▶ Play", command=self.player.play).pack(side="left", padx=2)
        ttk.Button(playback_frame, text="⏹ Stop", command=self.player.stop).pack(side="left", padx=2)
        ttk.Button(playback_frame, text="🔁 Toggle Loop", 
                  command=self.player.toggle_loop).pack(side="left", padx=2)
        
        # Tempo control
        tempo_frame = ttk.Frame(playback_frame)
        tempo_frame.pack(side="left", padx=5)
        ttk.Label(tempo_frame, text="Tempo:").pack(side="left")
        self.tempo_scale = ttk.Scale(tempo_frame, from_=0.25, to=2.0, 
                                   orient="horizontal", command=self.set_tempo)
        self.tempo_scale.set(1.0)
        self.tempo_scale.pack(side="left")
        
        # Right panel (Classification)
        right_panel = ttk.Frame(main_container)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # Classification buttons
        class_frame = ttk.LabelFrame(right_panel, text="Classification")
        class_frame.pack(fill="x", pady=5)
        
        ttk.Button(class_frame, text="OK (0)", 
                  command=lambda: self.classify("OK")).pack(fill="x", pady=2)
        
        ng_categories = [
            "1. Does not match the chord",
            "2. Melody is not good",
            "3. Rest is too long",
            "4. Too many small notes in a row",
            "5. Monotonous repetition",
            "6. Too much movement",
            "7. Same motif repeated",
            "8. Sounds like accompaniment"
        ]
        
        for i, cat in enumerate(ng_categories, 1):
            ttk.Button(class_frame, text=f"NG{i}: {cat} ({i})", 
                      command=lambda c=i: self.classify(f"NG{c}")
                      ).pack(fill="x", pady=2)
        
        # Comments section
        comments_frame = ttk.LabelFrame(right_panel, text="Comments")
        comments_frame.pack(fill="both", expand=True, pady=5)
        
        self.comments_text = scrolledtext.ScrolledText(comments_frame, height=4)
        self.comments_text.pack(fill="both", expand=True)
        
        # Statistics and progress
        stats_frame = ttk.LabelFrame(main_container, text="Statistics")
        stats_frame.pack(fill="x", pady=5)
        
        self.stats_label = ttk.Label(stats_frame, text="")
        self.stats_label.pack(pady=5)
        
        # Navigation and export
        nav_frame = ttk.Frame(main_container)
        nav_frame.pack(fill="x", pady=5)
        
        ttk.Button(nav_frame, text="⏮ Previous", 
                  command=self.prev_file).pack(side="left", padx=2)
        ttk.Button(nav_frame, text="⏭ Next", 
                  command=self.next_file).pack(side="left", padx=2)
        ttk.Button(nav_frame, text="📊 Export", 
                  command=self.export_results).pack(side="right", padx=2)
        
        self.update_stats()
        
    def setup_shortcuts(self):
        self.root.bind('0', lambda e: self.classify('OK'))
        for i in range(1, 9):
            self.root.bind(str(i), lambda e, i=i: self.classify(f'NG{i}'))
        self.root.bind('<space>', lambda e: self.toggle_playback())
        self.root.bind('<Left>', lambda e: self.prev_file())
        self.root.bind('<Right>', lambda e: self.next_file())
        
    def toggle_playback(self):
        if self.player.is_playing:
            self.player.stop()
        else:
            self.player.play()
            
    def set_tempo(self, value):
        self.player.set_tempo(float(value))
        
    def load_progress(self):
        try:
            with open('classification_progress.json', 'r') as f:
                data = json.load(f)
                self.classifications = data['classifications']
                self.stats.stats = data['stats']
                self.stats.total_time = timedelta(seconds=data['total_time_seconds'])
        except FileNotFoundError:
            pass
            
    def save_progress(self):
        data = {
            'classifications': self.classifications,
            'stats': self.stats.stats,
            'total_time_seconds': self.stats.total_time.total_seconds()
        }
        with open('classification_progress.json', 'w') as f:
            json.dump(data, f)
            
    def classify(self, classification: str):
        if not self.current_midi:
            return
            
        time_spent = datetime.now() - self.start_time if self.start_time else timedelta()
        
        entry = {
            'file': self.current_midi,
            'classification': classification,
            'comments': self.comments_text.get('1.0', 'end-1c'),
            'time_spent': str(time_spent),
            'timestamp': datetime.now().isoformat()
        }
        
        self.classifications.append(entry)
        self.stats.update(classification, time_spent)
        self.save_progress()
        self.update_stats()
        self.next_file()
        
    def update_stats(self):
        summary = self.stats.get_summary()
        stats_text = (
            f"Total files: {summary['total_files']}\n"
            f"OK ratio: {summary['ok_ratio']:.1%}\n"
            f"Most common NG: {summary['most_common_ng']}\n"
            f"Avg. time per file: {summary['avg_time_per_file'].seconds}s"
        )
        self.stats_label.config(text=stats_text)
        
    def load_file(self, filepath: str):
        try:
            self.player.load_midi(filepath)
            self.current_midi = filepath
            self.file_label.config(text=os.path.basename(filepath))
            self.comments_text.delete('1.0', 'end')
            self.start_time = datetime.now()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load MIDI file: {e}")
            
    def export_results(self):
        df = pd.DataFrame(self.classifications)
        filename = f'classifications_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
        
        # Export to multiple formats
        df.to_csv(f'{filename}.csv', index=False)
        df.to_excel(f'{filename}.xlsx', index=False)
        
        with open(f'{filename}.json', 'w') as f:
            json.dump(self.classifications, f, indent=2)
            
        messagebox.showinfo("Export Complete", 
                          f"Results exported to {filename}.csv/.xlsx/.json")

def main():
    root = tk.Tk()
    app = MIDIClassifierApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()