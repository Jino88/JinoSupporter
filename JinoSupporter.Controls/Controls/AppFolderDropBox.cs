using System;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;

namespace JinoSupporter.Controls
{
    /// <summary>
    /// 폴더 드래그앤드롭 + Browse 버튼을 제공하는 공통 폴더 입력 컨트롤.
    /// <see cref="AppFileDropBox"/>를 상속하며, Browse/Drop 동작만 폴더 선택으로 재정의.
    /// 폴더가 선택되면 <see cref="FolderSelected"/> 이벤트로 경로를 전달한다.
    /// </summary>
    public class AppFolderDropBox : AppFileDropBox
    {
        /// <summary>폴더가 선택됐을 때 발생.</summary>
        public event EventHandler<FolderSelectedEventArgs>? FolderSelected;

        public AppFolderDropBox()
        {
            // 기본 힌트 텍스트와 버튼 텍스트를 폴더용으로 변경
            HintText         = "Drag folder here";
            BrowseButtonText = "Browse Folder";
        }

        /// <summary>Browse 버튼 — 폴더 다이얼로그 오픈.</summary>
        protected override void OnBrowseClicked()
        {
            var dialog = new OpenFolderDialog
            {
                Title       = "Select folder",
                Multiselect = false
            };

            if (dialog.ShowDialog() == true)
                RaiseFolderSelected(dialog.FolderName);
        }

        /// <summary>드롭 — 드롭된 항목 중 첫 번째 폴더를 사용.</summary>
        protected override void OnDropped(string[] droppedPaths)
        {
            var folder = droppedPaths.FirstOrDefault(Directory.Exists);
            if (!string.IsNullOrWhiteSpace(folder))
                RaiseFolderSelected(folder);
        }

        private void RaiseFolderSelected(string folderPath)
        {
            FolderSelected?.Invoke(this, new FolderSelectedEventArgs(folderPath));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  EventArgs
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>폴더 선택 이벤트 데이터.</summary>
    public sealed class FolderSelectedEventArgs : EventArgs
    {
        /// <summary>선택된 폴더 경로.</summary>
        public string FolderPath { get; }

        public FolderSelectedEventArgs(string folderPath)
        {
            FolderPath = folderPath;
        }
    }
}
