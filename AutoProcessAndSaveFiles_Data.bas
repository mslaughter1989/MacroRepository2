Function LoadSFTPData() As Variant
    Dim sftpDataPart1 As Variant, sftpDataPart2 As Variant, sftpDataPart3 As Variant

    sftpDataPart1 = Array( _
        Array("PrismHR - MarvelHR (Pro PEO) (723827)", "MarvelHR_723827_mmddyyyy_Full.csv", "C:\Path\To\MarvelHR"), _
        Array("Clearway Energy", "CLEARWAY_622048_mmddyy_FULL.csv", "C:\Path\To\Clearway"), _
        Array("Aphena Pharma (349112, 349115, 570022)", "AphenaPharma_yyyymmdd.csv", "C:\Path\To\AphenaPharma") _
    )

    sftpDataPart2 = Array( _
        Array("Mightywell (with ARM)", "Recuro_MightyWELL_yyyymmdd.csv", "C:\Path\To\MightyWELL"), _
        Array("RS Utility Structures Inc (733441)", "RSUtilityStructuresInc_733441_mmddyyyy.csv.pgp.tmp", "C:\Path\To\RSUtility") _
    )

    sftpDataPart3 = Array( _
        Array("USA Hauling & Recycling", "USAHaulingRecycling_669506_mmddyyyy_Full.csv", "C:\Path\To\USAHauling") _
    )

    LoadSFTPData = CombineArrays(sftpDataPart1, sftpDataPart2, sftpDataPart3)
End Function
