"""Export via Office/LibreOffice and rasterize to local QA images."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


def export_pdf(source,output,soffice=None):
    pdf=output/(source.stem+'.pdf')
    if os.name=='nt' and not soffice:
        def quote(value):return "'"+str(value).replace("'","''")+"'"
        script='''$ErrorActionPreference='Stop'
$app=$null
$document=$null
try {
    foreach ($prog in @('kwps.Application','Word.Application')) {
        try {$app=New-Object -ComObject $prog;break} catch {}
    }
    if ($null -eq $app) {throw 'No WPS/Word automation available'}
    $app.Visible=$false
    $app.DisplayAlerts=0
    $document=$app.Documents.Open(SOURCE,$false,$true)
    $document.ExportAsFixedFormat(TARGET,17)
} finally {
    if ($null -ne $document) {$document.Close(0)}
    if ($null -ne $app) {$app.Quit()}
}
'''.replace('SOURCE',quote(source)).replace('TARGET',quote(pdf))
        with tempfile.TemporaryDirectory(prefix='cost-preview-') as folder:
            path=Path(folder)/'export.ps1';path.write_text(script,encoding='utf-8-sig')
            shell=shutil.which('pwsh.exe') or shutil.which('powershell.exe')
            if not shell:raise RuntimeError('PowerShell不可用，请指定--soffice')
            subprocess.run([shell,'-NoProfile','-NonInteractive','-File',str(path)],check=True,timeout=120)
    else:
        executable=soffice or shutil.which('soffice')
        if not executable:raise RuntimeError('未找到导出引擎，请指定--soffice或使用宿主文档渲染工具')
        with tempfile.TemporaryDirectory(prefix='cost-office-') as profile:
            subprocess.run([executable,'-env:UserInstallation='+Path(profile).as_uri(),'--headless',
                '--convert-to','pdf','--outdir',str(output),str(source)],check=True,timeout=120)
    if not pdf.is_file() or pdf.stat().st_size==0:raise RuntimeError('PDF导出未成功：'+str(pdf))
    return pdf


def main():
    p=argparse.ArgumentParser(description='仅生成本地排版核查预览')
    p.add_argument('document');p.add_argument('--output',required=True);p.add_argument('--soffice');p.add_argument('--pdftoppm')
    args=p.parse_args();source=Path(args.document).resolve();output=Path(args.output).resolve();output.mkdir(parents=True,exist_ok=True)
    raster=args.pdftoppm or shutil.which('pdftoppm')
    if not raster:raise RuntimeError('未找到pdftoppm，请安装Poppler或用--pdftoppm指定')
    # Clear only known preview images in this dedicated output folder; no recursion.
    for image in output.glob('page-*.png'):image.unlink()
    pdf=export_pdf(source,output,args.soffice)
    subprocess.run([raster,'-r','120','-png',str(pdf),str(output/'page')],check=True,timeout=120)
    images=list(output.glob('page-*.png'))
    if not images:raise RuntimeError('未生成预览图片')
    print(f'{len(images)}页预览：{output}')


if __name__=='__main__':main()
