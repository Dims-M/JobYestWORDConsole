using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MediaToolkit;
using MediaToolkit.Model;
using VideoLibrary;
using NReco.VideoConverter;


namespace JobYestWORDConsole
{

    /// <summary>
    /// Тестовой класс для работы с youtube
    /// </summary>
    public static class RabotaYoutube
    {
        

        public static void SaveMP3(string SaveToFolder, string VideoURL, string MP3Name)
        {
            string tempPath = @"C:\\1\\";
            try
            {
            var source = @SaveToFolder;
            var youtube = YouTube.Default;
            var vid = youtube.GetVideo(VideoURL);
         //  File.WriteAllBytes(@"C:\\1\\" + video.FullName, video.GetBytes());
           
           File.WriteAllBytes(source + vid.FullName, vid.GetBytes());
           //File.WriteAllBytes(@"C:\\1\\" + source + vid.FullName, vid.GetBytes());

            var inputFile = new MediaFile { Filename = source + vid.FullName };
            var outputFile = new MediaFile { Filename = $"{tempPath}{MP3Name}.mp3" };

            using (var engine = new Engine())
            {
                engine.GetMetadata(inputFile);

                engine.Convert(inputFile, outputFile);
                   
            }

            int a=1;

            }
            catch (Exception ex)
            {

            }
        }


        public static async Task SaveMp3V2Async()
        {

            using (var cli = Client.For(YouTube.Default))
            {
                string videoUrl1 = "https://www.youtube.com/watch?v=lzm5llVmR2E";
                string url = "YOUTUBE_URL";
                url = videoUrl1;

                var videoInfos = await cli.GetAllVideosAsync(url);
                //var possibleBitrates = videoInfos.Where(i => i.AdaptiveKind == AdaptiveKind.Audio).Select(i => i.AudioBitrate);
                //var possibleResolutions = videoInfos.Where(i => i.AdaptiveKind == AdaptiveKind.Video).Select(i => i.Resolution);
                foreach (var video in videoInfos)
                {
                    if (video.AdaptiveKind == AdaptiveKind.Audio)
                    //if(video.AudioFormat == AudioFormat.Aac)
                    //if(video.AudioBitrate == 128)
                    //if(video.AdaptiveKind == AdaptiveKind.Video)
                    //if (video.Format == VideoFormat.Mp4)
                    //if (video.Resolution == 360)
                    {
                        //Other methods
                    }
                }
                //OR
                //var downloadInfo = videoInfos.Where(i => i.AudioFormat == AudioFormat.Aac && i.AudioBitrate == 128).FirstOrDefault();
                var downloadInfo = videoInfos.Where(i => i.Format == VideoFormat.Mp4 && i.Resolution == 720).FirstOrDefault(); // if 720p is possible
                string downloadUri = downloadInfo.Uri;
               
            }

        }
            //public static async Task TestUtAsync()
            //{
            //    var client = new YoutubeClient();
            //    var videoId = NormalizeVideoId(txtFileURL.Text);
            //    var video = await client.GetVideoAsync(videoId);
            //    var streamInfoSet = await client.GetVideoMediaStreamInfosAsync(videoId);
            //    // Get the best muxed stream
            //    var streamInfo = streamInfoSet.Muxed.WithHighestVideoQuality();
            //    // Compose file name, based on metadata
            //    var fileExtension = streamInfo.Container.GetFileExtension();
            //    var fileName = $"{video.Title}.{fileExtension}";
            //    // Replace illegal characters in file name

            //    fileName = RemoveIllegalFileNameChars(fileName);
            //    tmrVideo.Enabled = true;

            //    // Download video
            //    txtMessages.Text = "Downloading Video please wait ... ";

            //    //using (var progress = new ProgressBar())
            //    await client.DownloadMediaStreamAsync(streamInfo, fileName);

            //    // Add Nuget package: https://www.nuget.org/packages/NReco.VideoConverter/ To Convert MP4 to MP3
            //    if (ckbAudioOnly.Checked)
            //    {
            //        var Convert = new NReco.VideoConverter.FFMpegConverter();
            //        String SaveMP3File = MP3FolderPath + fileName.Replace(".mp4", ".mp3");

            //        Convert.ConvertMedia(fileName, SaveMP3File, "mp3");
            //        //Delete the MP4 file after conversion
            //        File.Delete(fileName);
            //        LoadMP3Files();
            //        txtMessages.Text = "File Converted to MP3";
            //        tmrVideo.Enabled = false;
            //        txtMessages.BackColor = Color.White;
            //        if (ckbAutoPlay.Checked) { PlayFile(SaveMP3File); }
            //        return;
            //    }
            //}
        
    }
}
