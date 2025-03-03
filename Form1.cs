using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MediaDevices;
using System.IO.Compression;
using System.Collections.Generic;

namespace DaisyBookManager
{
    public partial class Form1 : Form
    {
        private ListBox daisyListBox = null!;
        private Button selectFolderButton = null!;
        private Button copyButton = null!;
        private Button refreshButton = null!;
        private Label statusLabel = null!;
        private string targetFolder = "";
        private MediaDevice? selectedDevice;
        private readonly string _tempFolder = Path.Combine(
            Path.GetTempPath(), 
            "DaisyBookManager");
        private readonly string _appDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DaisyBookManager");
        private readonly string _settingsFile;
        private CancellationTokenSource? _cancellationTokenSource;

        public Form1()
        {
            InitializeComponent();
            _settingsFile = Path.Combine(_appDataFolder, "settings.cfg");
            EnsureAppDataFolderExists();
            EnsureTempFolderExists();
            SetupCustomUI();
            LoadSettings();
            
            // Aseta ensin latausviesti näkyviin
            daisyListBox.Items.Add("Etsitään DAISY-kirjoja...");
            
            // Käynnistä kirjojen lataus viiveellä, jotta käyttöliittymä ehtii alustua
            System.Windows.Forms.Timer startupTimer = new System.Windows.Forms.Timer();
            startupTimer.Interval = 100;
            startupTimer.Tick += (s, e) => {
                startupTimer.Stop();
                startupTimer.Dispose();
                LoadDaisyBooks();
            };
            startupTimer.Start();
        }

        private void SetupCustomUI()
        {
            this.Text = "Stream 3 Apuri";
            this.Width = 600;
            this.Height = 500;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            statusLabel = new Label
            {
                Text = "Valitse ladattu DAISY-kirja ja kopioi se Victor Stream -laitteeseen.",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                TabIndex = 4
            };

            daisyListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.One,
                Font = new Font("Segoe UI", 10F),
                TabIndex = 0
            };
            daisyListBox.SelectedIndexChanged += DaisyListBox_SelectedIndexChanged;

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 120
            };

            refreshButton = new Button
            {
                Text = "Päivitä kirjaluettelo",
                Width = 180,
                Height = 30,
                Location = new Point(20, 20),
                TabIndex = 1
            };
            refreshButton.Click += RefreshButton_Click;

            selectFolderButton = new Button
            {
                Text = "Valitse Stream-kansio",
                Width = 180,
                Height = 30,
                Location = new Point(20, 60),
                TabIndex = 2
            };
            selectFolderButton.Click += async (s, e) => await SelectTargetFolderAsync();

            copyButton = new Button
            {
                Text = "Kopioi valittu kirja",
                Width = 180,
                Height = 30,
                Location = new Point(220, 20),
                Enabled = false,
                TabIndex = 3
            };
            copyButton.Click += async (s, e) => await CopySelectedBookAsync();

            buttonPanel.Controls.Add(copyButton);
            buttonPanel.Controls.Add(selectFolderButton);
            buttonPanel.Controls.Add(refreshButton);

            Controls.Add(statusLabel);
            Controls.Add(buttonPanel);
            Controls.Add(daisyListBox);
        }

        private void EnsureAppDataFolderExists()
        {
            try
            {
                if (!Directory.Exists(_appDataFolder))
                {
                    Directory.CreateDirectory(_appDataFolder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe asetuskansion luomisessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureTempFolderExists()
        {
            try
            {
                if (Directory.Exists(_tempFolder))
                {
                    Directory.Delete(_tempFolder, true);
                }
                
                Directory.CreateDirectory(_tempFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe väliaikaiskansion luomisessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CleanupTempFolder()
        {
            try
            {
                if (Directory.Exists(_tempFolder))
                {
                    Directory.Delete(_tempFolder, true);
                    Directory.CreateDirectory(_tempFolder);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Virhe väliaikaiskansion siivoamisessa: {ex.Message}");
            }
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFile))
                {
                    string[] settings = File.ReadAllLines(_settingsFile);
                    
                    if (settings.Length >= 2 && !string.IsNullOrEmpty(settings[0]) && !string.IsNullOrEmpty(settings[1]))
                    {
                        string deviceName = settings[0];
                        string folderPath = settings[1];
                        
                        // Yritetään löytää ja yhdistää laitteeseen
                        var devices = MediaDevice.GetDevices();
                        var matchingDevice = devices.FirstOrDefault(d => d.FriendlyName == deviceName);
                        
                        if (matchingDevice != null)
                        {
                            try
                            {
                                selectedDevice = matchingDevice;
                                selectedDevice.Connect();
                                
                                // Tarkistetaan että kansio on olemassa
                                selectedDevice.GetDirectories(folderPath);
                                targetFolder = folderPath;
                                statusLabel.Text = $"Kohdekansio: {targetFolder}";
                            }
                            catch
                            {
                                // Jos kansiota ei löydy tai yhdistäminen epäonnistuu, nollataan laitetiedot
                                DisconnectDevice();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe asetusten lataamisessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettings()
        {
            try
            {
                // Luodaan lista tiedostoon tallennettavista riveistä
                List<string> settings = new List<string>();
                
                // 1. Laitteen nimi - käytetään aina "Stream V3"
                settings.Add("Stream V3");
                
                // 2. Kohdekansio
                settings.Add(targetFolder ?? "");
                
                // Kirjoitetaan tiedostoon
                File.WriteAllLines(_settingsFile, settings);
                
                Console.WriteLine("Asetukset tallennettu:");
                Console.WriteLine($"Laite: Stream V3");
                Console.WriteLine($"Kansio: {targetFolder}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe asetusten tallentamisessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DaisyListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Enable copy button only if a book is selected and a target folder is chosen
            if (daisyListBox.SelectedItem != null && daisyListBox.SelectedItem is DaisyBookInfo)
            {
                copyButton.Enabled = !string.IsNullOrEmpty(targetFolder);
            }
            else
            {
                copyButton.Enabled = false;
            }
        }

        private async void RefreshButton_Click(object sender, EventArgs e)
        {
            // Disable UI during search
            EnableUI(false);
            
            try
            {
                // Näytä latausviesti ruudunlukijalle
                statusLabel.Text = "Etsitään DAISY-kirjoja, odota hetki...";
                statusLabel.Focus(); // Anna hetkellinen fokus tilaviestille ruudunlukijaa varten
                Application.DoEvents(); // Varmista että viesti näytetään
                
                await LoadDaisyBooksAsync();
            }
            finally
            {
                // Re-enable UI when done
                EnableUI(true);
                
                // Varmista että fokus palaa takaisin listaan
                if (daisyListBox.Items.Count > 0)
                {
                    daisyListBox.Focus();
                    // Ilmoitus ruudunlukijalle että lataus on valmis
                    statusLabel.Text = $"Kirjahylly päivitetty. {statusLabel.Text}";
                }
            }
        }

        private void EnableUI(bool enable)
        {
            refreshButton.Enabled = enable;
            selectFolderButton.Enabled = enable;
            copyButton.Enabled = enable && daisyListBox.SelectedItem != null && 
                                daisyListBox.SelectedItem is DaisyBookInfo && 
                                !string.IsNullOrEmpty(targetFolder);
            daisyListBox.Enabled = enable;
        }

        private void LoadDaisyBooks()
        {
            // Käynnistä asynkroninen lataus automaattisesti
            Task.Run(async () => await LoadDaisyBooksAsync()).ConfigureAwait(false);
        }

        private async Task LoadDaisyBooksAsync()
        {
            string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

            if (!Directory.Exists(downloadsFolder))
            {
                MessageBox.Show("Latauskansiota ei löytynyt.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var zipFiles = Directory.GetFiles(downloadsFolder, "*.zip").ToList();
            
            daisyListBox.Items.Clear();
            
            if (zipFiles.Count == 0)
            {
                daisyListBox.Items.Add("Ei löydetty ZIP-tiedostoja.");
                copyButton.Enabled = false;
                return;
            }
            
            // Placeholder message during processing
            daisyListBox.Items.Add("Etsitään DAISY-kirjoja...");
            statusLabel.Text = $"Käsitellään {zipFiles.Count} ZIP-tiedostoa...";
            daisyListBox.Refresh();
            
            // Find DAISY books in background
            List<DaisyBookInfo> daisyBooks = await Task.Run(() => {
                var books = new List<DaisyBookInfo>();
                
                // Process each ZIP file
                foreach (var zipFile in zipFiles)
                {
                    try
                    {
                        var bookInfo = GetDaisyBookInfo(zipFile);
                        if (bookInfo != null)
                        {
                            books.Add(bookInfo);
                        }
                    }
                    catch
                    {
                        // Skip invalid files
                    }
                }
                
                return books;
            });
            
            // Update UI with results
            daisyListBox.Items.Clear();
            if (daisyBooks.Count == 0)
            {
                daisyListBox.Items.Add("Ei löydetty DAISY-kirjoja.");
                copyButton.Enabled = false;
            }
            else
            {
                foreach (var book in daisyBooks)
                {
                    daisyListBox.Items.Add(book);
                }
                statusLabel.Text = $"Löydetty {daisyBooks.Count} DAISY-kirjaa latauskansiosta.";
            }
            
            // Remove temporary extraction folders
            CleanupTempFolder();
        }

        private DaisyBookInfo GetDaisyBookInfo(string zipFilePath)
        {
            // Create unique folder for this book
            string extractPath = Path.Combine(_tempFolder, Path.GetFileNameWithoutExtension(zipFilePath));
            
            try
            {
                // Create a unique subdirectory for testing
                string testPath = Path.Combine(extractPath, "test");
                Directory.CreateDirectory(testPath);
                
                // Try to find ncc.html without extracting all files
                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    // Look for ncc.html or NCC.HTML
                    var nccEntry = archive.Entries.FirstOrDefault(e => 
                        e.Name.Equals("ncc.html", StringComparison.OrdinalIgnoreCase));
                    
                    if (nccEntry == null)
                    {
                        return null; // Not a DAISY book
                    }
                    
                    // Extract only the ncc.html file
                    nccEntry.ExtractToFile(Path.Combine(testPath, nccEntry.Name), true);
                    
                    // Read the NCC file
                    string nccContent = File.ReadAllText(Path.Combine(testPath, nccEntry.Name));
                    
                    // Extract metadata
                    string title = ExtractMetadata(nccContent, "dc:title");
                    string author = ExtractMetadata(nccContent, "dc:creator");
                    
                    // Use filename as fallback
                    if (string.IsNullOrEmpty(title))
                    {
                        title = Path.GetFileNameWithoutExtension(zipFilePath);
                    }
                    
                    if (string.IsNullOrEmpty(author))
                    {
                        author = "Tuntematon";
                    }
                    
                    // Return book info
                    return new DaisyBookInfo
                    {
                        ZipPath = zipFilePath,
                        Title = title,
                        Author = author,
                        FolderName = GetSafeFileName(title)
                    };
                }
            }
            catch
            {
                return null; // Not a valid ZIP or error in processing
            }
            finally
            {
                // Clean up the test directory
                try
                {
                    if (Directory.Exists(extractPath))
                    {
                        Directory.Delete(extractPath, true);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
        
        private string ExtractMetadata(string html, string metadataName)
        {
            // Simple string-based parsing
            string startTag = $"<meta name=\"{metadataName}\" content=\"";
            int startIndex = html.IndexOf(startTag);
            
            if (startIndex == -1)
                return string.Empty;
                
            startIndex += startTag.Length;
            int endIndex = html.IndexOf("\"", startIndex);
            
            if (endIndex == -1)
                return string.Empty;
                
            return html.Substring(startIndex, endIndex - startIndex).Trim();
        }
        
        private string GetSafeFileName(string input)
        {
            // Replace invalid characters
            string safeName = string.Join("_", input.Split(Path.GetInvalidFileNameChars()));
            
            // Limit length
            if (safeName.Length > 50)
            {
                safeName = safeName.Substring(0, 50);
            }
            
            return safeName;
        }

        private async Task SelectTargetFolderAsync()
        {
            await Task.Run(() => DisconnectDevice()); // Ensure any previous device is disconnected

            var devices = MediaDevice.GetDevices().ToList();
            if (devices.Count == 0)
            {
                MessageBox.Show("Ei löydetty yhtään WPD-laitetta. Liitä Victor Stream ja yritä uudelleen.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Allow user to select from multiple devices if there are more than one
                if (devices.Count > 1)
                {
                    var deviceNames = devices.Select(d => d.FriendlyName).ToArray();
                    string selectedDeviceName = PromptForSelection("Valitse laite", deviceNames);
                    
                    if (string.IsNullOrEmpty(selectedDeviceName))
                    {
                        return; // User cancelled
                    }
                    
                    selectedDevice = devices.FirstOrDefault(d => d.FriendlyName == selectedDeviceName);
                }
                else
                {
                    selectedDevice = devices.FirstOrDefault();
                }

                if (selectedDevice == null)
                {
                    MessageBox.Show("Laitetta ei pystytty valitsemaan.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                await Task.Run(() => {
                    selectedDevice.Connect();
                });
                
                statusLabel.Text = $"Yhdistetty laitteeseen: {selectedDevice.FriendlyName}";

                // Get root directories and browse for folder
                var rootPath = @"\";
                await BrowseForFolderAsync(rootPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe laitteeseen yhdistämisessä: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task BrowseForFolderAsync(string currentPath)
        {
            try
            {
                // Get directories in background to avoid UI freezing
                List<string> directories = await Task.Run(() => {
                    return selectedDevice.GetDirectories(currentPath).ToList();
                });

                if (directories.Count == 0)
                {
                    // If there are no subdirectories, use this folder
                    targetFolder = currentPath;
                    copyButton.Enabled = daisyListBox.SelectedItem != null && 
                                        daisyListBox.SelectedItem is DaisyBookInfo;
                    MessageBox.Show($"Kohdekansio valittu: {targetFolder}", "Valinta onnistui", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = $"Kohdekansio: {targetFolder}";
                    
                    // Save settings
                    SaveSettings();
                    return;
                }

                // Add option to use current folder
                var options = new string[] { "[Käytä tätä kansiota]" }.Concat(directories.Select(Path.GetFileName)).ToArray();
                
                string selection = PromptForSelection($"Valitse kansio: {currentPath}", options);
                
                if (string.IsNullOrEmpty(selection))
                {
                    return; // User cancelled
                }
                
                if (selection == "[Käytä tätä kansiota]")
                {
                    targetFolder = currentPath;
                    copyButton.Enabled = daisyListBox.SelectedItem != null && 
                                        daisyListBox.SelectedItem is DaisyBookInfo;
                    MessageBox.Show($"Kohdekansio valittu: {targetFolder}", "Valinta onnistui", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    statusLabel.Text = $"Kohdekansio: {targetFolder}";
                    
                    // Save settings
                    SaveSettings();
                }
                else
                {
                    // Navigate to selected subfolder
                    string selectedPath = Path.Combine(currentPath, selection).Replace('\\', '/');
                    await BrowseForFolderAsync(selectedPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe kansioiden selaamisessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string PromptForSelection(string prompt, string[] options)
        {
            using (var form = new Form())
            {
                form.Text = prompt;
                form.StartPosition = FormStartPosition.CenterParent;
                form.Width = 400;
                form.Height = 300;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var listBox = new ListBox
                {
                    Dock = DockStyle.Fill,
                    SelectionMode = SelectionMode.One,
                    Font = new Font("Segoe UI", 10F)
                };
                listBox.Items.AddRange(options);
                
                if (listBox.Items.Count > 0)
                {
                    listBox.SelectedIndex = 0;
                }

                var okButton = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Dock = DockStyle.Bottom,
                    Height = 30
                };

                var cancelButton = new Button
                {
                    Text = "Peruuta",
                    DialogResult = DialogResult.Cancel,
                    Dock = DockStyle.Bottom,
                    Height = 30
                };

                form.Controls.Add(listBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                DialogResult result = form.ShowDialog();
                if (result == DialogResult.OK && listBox.SelectedItem != null)
                {
                    return listBox.SelectedItem.ToString();
                }
                
                return null; // User cancelled
            }
        }

        private async Task CopySelectedBookAsync()
        {
            if (daisyListBox.SelectedItem == null || !(daisyListBox.SelectedItem is DaisyBookInfo))
            {
                MessageBox.Show("Valitse ensin DAISY-kirja luettelosta.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(targetFolder) || selectedDevice == null || !selectedDevice.IsConnected)
            {
                MessageBox.Show("Valitse ensin Victor Streamin kohdekansio.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get the selected book
            var selectedBook = (DaisyBookInfo)daisyListBox.SelectedItem;
            
            // Check if source file exists
            if (!File.Exists(selectedBook.ZipPath))
            {
                MessageBox.Show("Lähdetiedostoa ei löytynyt. Tarkista, että DAISY-kirja on olemassa.", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Prepare destination path
            string destinationPath = targetFolder.Replace('\\', '/');
            if (!destinationPath.EndsWith("/"))
            {
                destinationPath += "/";
            }
            string destinationFolder = destinationPath + selectedBook.FolderName;

            // Check if folder already exists on the device
            bool folderExists = false;
            try
            {
                var existingFolders = await Task.Run(() => selectedDevice.GetDirectories(destinationPath));
                folderExists = existingFolders.Any(f => Path.GetFileName(f) == selectedBook.FolderName);
                
                if (folderExists)
                {
                    var result = MessageBox.Show(
                        $"Kansio {selectedBook.FolderName} on jo olemassa kohdelaitteella. Haluatko korvata sen?",
                        "Kansio on jo olemassa",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                        
                    if (result == DialogResult.No)
                    {
                        return;
                    }
                    
                    // Delete existing folder
                    await Task.Run(() => selectedDevice.DeleteDirectory(destinationFolder, true));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe kansion tarkistuksessa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Show copy progress dialog
            using (var progressForm = new ImprovedCopyProgressForm())
            {
                // Disable UI during copy
                EnableUI(false);
                
                try
                {
                    // Start the copy operation
                    var result = await progressForm.CopyDaisyBookAsync(selectedDevice, selectedBook, destinationFolder);
                    
                    if (result)
                    {
                        statusLabel.Text = $"Kirja {selectedBook.Title} kopioitu onnistuneesti!";
                    }
                    else
                    {
                        statusLabel.Text = $"Kopiointi keskeytyi tai epäonnistui.";
                    }
                }
                finally
                {
                    // Re-enable UI
                    EnableUI(true);
                    
                    // Clean up temp folder
                    CleanupTempFolder();
                }
            }
        }

        private void DisconnectDevice()
        {
            if (selectedDevice != null && selectedDevice.IsConnected)
            {
                try
                {
                    selectedDevice.Disconnect();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Virhe laitteen irrottamisessa: {ex.Message}");
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            
            // Cancel any ongoing operations
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
            }
            
            // Disconnect device when closing the form
            DisconnectDevice();
            
            // Clean up temp folder
            CleanupTempFolder();
        }
    }

    public class DaisyBookInfo
    {
        public string ZipPath { get; set; } = string.Empty;
        public string ExtractPath { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"{Title} - {Author}";
        }
    }

    public class ImprovedCopyProgressForm : Form
    {
        private ProgressBar progressBar = null!;
        private Label statusLabel = null!;
        private Button cancelButton = null!;
        private CancellationTokenSource? _cancellationTokenSource;
        private long _totalSize;
        private long _copiedBytes;
        private string _tempFolder;
        
        public ImprovedCopyProgressForm()
        {
            InitializeUI();
            
            // Create temp folder path
            _tempFolder = Path.Combine(
                Path.GetTempPath(), 
                "DaisyBookManager", 
                Guid.NewGuid().ToString());
        }

        private void InitializeUI()
        {
            Text = "Kopiointi käynnissä";
            Width = 500;
            Height = 180;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;

            statusLabel = new Label
            {
                Text = "Valmistellaan kopiointia...",
                Dock = DockStyle.Top,
                Height = 40,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10F)
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 30,
                Minimum = 0,
                Maximum = 100,
                Style = ProgressBarStyle.Continuous
            };
            
            cancelButton = new Button
            {
                Text = "Peruuta",
                Dock = DockStyle.Bottom,
                Height = 30
            };
            cancelButton.Click += (s, e) => {
                if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
                {
                    _cancellationTokenSource.Cancel();
                    statusLabel.Text = "Peruutetaan...";
                    cancelButton.Enabled = false;
                }
            };

            Controls.Add(cancelButton);
            Controls.Add(progressBar);
            Controls.Add(statusLabel);
        }

        public async Task<bool> CopyDaisyBookAsync(MediaDevice device, DaisyBookInfo book, string destinationFolder)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _copiedBytes = 0;
            
            try
            {
                // Show the form
                this.Show();
                Application.DoEvents();
                
                // Step 1: Extract ZIP to temporary folder
                statusLabel.Text = "Puretaan ZIP-tiedostoa...";
                progressBar.Value = 0;
                
                // Create temp directory for extraction
                Directory.CreateDirectory(_tempFolder);
                book.ExtractPath = _tempFolder;
                
                // Extract in background
                await Task.Run(() => {
                    try
                    {
                        ZipFile.ExtractToDirectory(book.ZipPath, _tempFolder);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Virhe ZIP-tiedoston purkamisessa: {ex.Message}", ex);
                    }
                });
                
                if (_cancellationTokenSource.IsCancellationRequested)
                    return false;
                
                // Step 2: Calculate total size for progress tracking
                statusLabel.Text = "Lasketaan tiedostojen kokoa...";
                progressBar.Value = 5;
                
                _totalSize = await Task.Run(() => GetDirectorySize(_tempFolder));
                
                if (_cancellationTokenSource.IsCancellationRequested)
                    return false;
                
                // Step 3: Create target folder
                statusLabel.Text = "Luodaan kohdekansiota...";
                progressBar.Value = 10;
                
                await Task.Run(() => {
                    try
                    {
                        device.CreateDirectory(destinationFolder);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Virhe kohdekansion luomisessa: {ex.Message}", ex);
                    }
                });
                
                if (_cancellationTokenSource.IsCancellationRequested)
                    return false;
                
                // Step 4: Copy files
                statusLabel.Text = "Kopioidaan tiedostoja...";
                progressBar.Value = 15;
                
                bool success = await CopyDaisyFolderAsync(device, _tempFolder, destinationFolder);
                
                if (success && !_cancellationTokenSource.IsCancellationRequested)
                {
                    progressBar.Value = 100;
                    statusLabel.Text = "Kopiointi valmis!";
                    MessageBox.Show("Kirja kopioitu onnistuneesti!", "Valmis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                else
                {
                    if (_cancellationTokenSource.IsCancellationRequested)
                    {
                        statusLabel.Text = "Kopiointi peruutettu.";
                    }
                    else
                    {
                        statusLabel.Text = "Kopiointi epäonnistui.";
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Virhe kopioinnissa: {ex.Message}", "Virhe", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                // Clean up
                try
                {
                    if (Directory.Exists(_tempFolder))
                    {
                        Directory.Delete(_tempFolder, true);
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
        
        private async Task<bool> CopyDaisyFolderAsync(MediaDevice device, string sourceDirectory, string destinationDirectory)
        {
            // Create device directory if it doesn't exist
            if (!device.DirectoryExists(destinationDirectory))
            {
                device.CreateDirectory(destinationDirectory);
            }

            // Get all files and directories
            var files = Directory.GetFiles(sourceDirectory);
            var directories = Directory.GetDirectories(sourceDirectory);
            
            // First copy all files
            int fileNumber = 0;
            int totalFiles = files.Length;
            
            foreach (var file in files)
            {
                if (_cancellationTokenSource.IsCancellationRequested)
                    return false;
                
                fileNumber++;
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(destinationDirectory, fileName).Replace('\\', '/');
                
                // Update UI with current file
                UpdateProgress($"Kopioidaan ({fileNumber}/{totalFiles}): {fileName}", fileNumber, totalFiles);
                
                // Copy in background thread
                await Task.Run(() => {
                    try
                    {
                        device.UploadFile(file, destFile);
                        
                        // Update copied bytes for progress calculation
                        Interlocked.Add(ref _copiedBytes, new FileInfo(file).Length);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Virhe tiedoston kopioinnissa: {ex.Message}", ex);
                    }
                });
                
                // Update progress based on total bytes
                UpdateByteProgress();
            }
            
            // Then copy all subdirectories recursively
            foreach (var directory in directories)
            {
                if (_cancellationTokenSource.IsCancellationRequested)
                    return false;
                
                string dirName = Path.GetFileName(directory);
                string destDir = Path.Combine(destinationDirectory, dirName).Replace('\\', '/');
                
                // Recursively copy subdirectories
                bool success = await CopyDaisyFolderAsync(device, directory, destDir);
                if (!success)
                    return false;
            }
            
            return true;
        }
        
        private void UpdateProgress(string status, int current, int total)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => UpdateProgress(status, current, total)));
                return;
            }
            
            // Calculate percentage (limit to files part - 15-80%)
            int percentage = 15 + (current * 65) / Math.Max(1, total);
            statusLabel.Text = status;
            progressBar.Value = Math.Min(80, percentage); // Cap at 80% for files
        }
        
        private void UpdateByteProgress()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(UpdateByteProgress));
                return;
            }
            
            if (_totalSize > 0)
            {
                int percentage = (int)((_copiedBytes * 100) / _totalSize);
                progressBar.Value = Math.Min(99, percentage); // Cap at 99% until complete
                
                string copiedMB = (_copiedBytes / (1024.0 * 1024.0)).ToString("F1");
                string totalMB = (_totalSize / (1024.0 * 1024.0)).ToString("F1");
                statusLabel.Text = $"Kopioitu: {copiedMB} Mt / {totalMB} Mt ({percentage}%)";
            }
        }
        
        private long GetDirectorySize(string path)
        {
            long size = 0;
            
            // Add up sizes of all files
            foreach (string file in Directory.GetFiles(path, "*", SearchOption.AllDirectories))
            {
                size += new FileInfo(file).Length;
            }
            
            return size;
        }
    }
}
