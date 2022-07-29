using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace Excel {
    
    public class Interpreter : IDisposable {

        private Process _interpreter;
        
        public Interpreter(string relativeFilePath) {
            var interpreterSettings = new ProcessStartInfo {
                CreateNoWindow = true,
                UseShellExecute = false,
                FileName =  Path.GetFullPath(@"..\..\" + relativeFilePath),
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardInput = true,
                RedirectStandardOutput = true
            };

            _interpreter = new Process();
            _interpreter.StartInfo = interpreterSettings;
            _interpreter.Start();
        }

        public string Evaluate(string cell, string expression) {
            _interpreter.StandardInput.WriteLine("echo ${" + cell + "} = " + expression);
            return _interpreter.StandardOutput.ReadLine();
        }

        public void Dispose() {
            _interpreter?.Kill();
            _interpreter?.Dispose();
        }
        
    }
}