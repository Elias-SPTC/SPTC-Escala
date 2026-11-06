package com.example.sptc_escala

import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.gestures.detectHorizontalDragGestures
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.layout.ExperimentalLayoutApi
import androidx.compose.foundation.layout.FlowRow
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.pager.HorizontalPager
import androidx.compose.foundation.pager.rememberPagerState
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.CircleShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.ChevronLeft
import androidx.compose.material.icons.filled.ChevronRight
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.Engineering
import androidx.compose.material.icons.filled.FileOpen
import androidx.compose.material.icons.filled.FilterList
import androidx.compose.material.icons.filled.Nightlight
import androidx.compose.material.icons.filled.PersonSearch
import androidx.compose.material.icons.filled.WbSunny
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.lifecycle.ViewModel
import androidx.lifecycle.ViewModelProvider
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewModelScope
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.room.*
import com.example.sptc_escala.ui.theme.SPTCEscalaTheme
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.*
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.InputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

// ==========================================
// 1. MODELOS DE DADOS (ROOM ENTITIES)
// ==========================================

@Entity(tableName = "employees")
data class EmployeeEntity(
    @PrimaryKey val id: Int,
    val name: String,
    val role: String, // "PERITO" ou "FOTOGRAFO"
)

@Entity(
    tableName = "shifts",
    indices = [
        Index("date"),
        Index("plantaoExpertId"),
        Index("plantaoPhotographerId"),
        Index("engenhariaExpertId"),
        Index("engenhariaPhotographerId"),
    ]
)
data class ShiftEntity(
    @PrimaryKey(autoGenerate = true) val id: Long = 0,
    val date: String,            // Formato DD/MM/YYYY
    val shiftNumber: Int,        // 1 (Diurno) ou 2 (Noturno)
    val plantaoExpertId: Int?,
    val plantaoPhotographerId: Int?,
    val engenhariaExpertId: Int?,
    val engenhariaPhotographerId: Int?
)

data class ShiftWithDetails(
    @Embedded val shift: ShiftEntity,
    @Relation(parentColumn = "plantaoExpertId", entityColumn = "id")
    val plantaoExpert: EmployeeEntity?,
    @Relation(parentColumn = "plantaoPhotographerId", entityColumn = "id")
    val plantaoPhotographer: EmployeeEntity?,
    @Relation(parentColumn = "engenhariaExpertId", entityColumn = "id")
    val engenhariaExpert: EmployeeEntity?,
    @Relation(parentColumn = "engenhariaPhotographerId", entityColumn = "id")
    val engenhariaPhotographer: EmployeeEntity?
)

// ==========================================
// 2. DAOS E BANCO DE DADOS
// ==========================================

@Dao
interface EmployeeDao {
    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertAll(employees: List<EmployeeEntity>)

    @Query("DELETE FROM employees")
    suspend fun clearEmployees()

    @Query("SELECT * FROM employees ORDER BY name ASC")
    fun getAllEmployeesFlow(): Flow<List<EmployeeEntity>>

    @Query("SELECT * FROM employees")
    suspend fun getAll(): List<EmployeeEntity>
}

@Dao
interface ShiftDao {
    @Transaction
    @Query("SELECT * FROM shifts ORDER BY date ASC, shiftNumber ASC")
    fun getAllShiftsWithDetails(): Flow<List<ShiftWithDetails>>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertShifts(shifts: List<ShiftEntity>)

    @Query("SELECT * FROM shifts WHERE date = :date AND shiftNumber = :shiftNumber LIMIT 1")
    suspend fun getShift(date: String, shiftNumber: Int): ShiftEntity?

    @Query("UPDATE shifts SET engenhariaExpertId = :expertId WHERE date = :date AND shiftNumber = :shiftNumber")
    suspend fun updateEngenhariaExpert(date: String, shiftNumber: Int, expertId: Int?)

    @Query("UPDATE shifts SET plantaoExpertId = :pExpId, plantaoPhotographerId = :pFotId, engenhariaExpertId = :eExpId, engenhariaPhotographerId = :eFotId WHERE id = :shiftId")
    suspend fun updateShiftAssignments(shiftId: Long, pExpId: Int?, pFotId: Int?, eExpId: Int?, eFotId: Int?)

    @Query("DELETE FROM shifts WHERE date LIKE '%/' || :monthStr || '/' || :yearStr")
    suspend fun clearShiftsForMonth(monthStr: String, yearStr: String)

    @Query("DELETE FROM shifts")
    suspend fun clearShifts()
}

@Database(entities = [EmployeeEntity::class, ShiftEntity::class], version = 1, exportSchema = false)
abstract class AppDatabase : RoomDatabase() {
    abstract fun employeeDao(): EmployeeDao
    abstract fun shiftDao(): ShiftDao

    companion object {
        @Volatile
        private var INSTANCE: AppDatabase? = null

        fun getDatabase(context: Context): AppDatabase {
            return INSTANCE ?: synchronized(this) {
                val instance = Room.databaseBuilder(
                    context.applicationContext,
                    AppDatabase::class.java,
                    "escala_plantao_database"
                )
                .fallbackToDestructiveMigration(dropAllTables = true)
                .build()
                INSTANCE = instance
                instance
            }
        }
    }
}

// ==========================================
// 3. LEITOR / IMPORTADOR EXCEL RESILIENTE
// ==========================================

class ScheduleExcelImporter(
    private val context: Context,
    private val employeeDao: EmployeeDao,
    private val shiftDao: ShiftDao
) {
    suspend fun importExcelToRoom(uri: Uri): Result<String> = runCatching {
        try {
            context.contentResolver.takePersistableUriPermission(
                uri,
                Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        } catch (_: Exception) {}

        val inputStream: InputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalArgumentException("Não foi possível abrir o arquivo selecionado.")

        val formatter = DataFormatter()

        inputStream.use { stream ->
            val workbook = WorkbookFactory.create(stream)
            workbook.use { wb ->
                if (wb.numberOfSheets == 0) {
                    throw IllegalArgumentException("A planilha está vazia.")
                }

                val sheet = wb.getSheetAt(0) // Planilha principal da SPTC

                // 1. Extrair Mês e Ano do cabeçalho
                var monthStr = "10"
                var yearStr = "2026"

                for (rowIndex in 0..12) {
                    val row = sheet.getRow(rowIndex) ?: continue
                    for (cell in row) {
                        val text = formatter.formatCellValue(cell).uppercase()
                        if (text.contains("ESCALA DE PLANTÃO") || text.contains("OUTUBRO") || text.contains("202")) {
                            val extracted = parseMonthAndYear(text)
                            if (extracted != null) {
                                monthStr = extracted.first
                                yearStr = extracted.second
                                break
                            }
                        }
                    }
                }

                // 2. Mapear Funcionários das tabelas secundárias à direita (Peritos e Fotógrafos)
                val peritosMapRaw = scanSideTable(sheet, formatter, "PERITO")
                val fotografoMapRaw = scanSideTable(sheet, formatter, "FOTOGRAFO")

                // Fallback de segurança para a escala de Sorocaba caso a leitura de células mescladas no Excel não encontre os objetos
                val defaultPeritos = mapOf(
                    1 to "Amanda", 2 to "Andrea", 3 to "Artur",
                    4 to "Guilherme M.", 5 to "Laura", 6 to "Paloma", 7 to "Renan"
                )
                val defaultFotografos = mapOf(
                    1 to "Antonio", 2 to "Baena", 3 to "Elias", 4 to "Gilberto",
                    5 to "Lucas", 6 to "Delmont", 7 to "Maurício S", 8 to "Pecchia", 9 to "Thiago"
                )

                val peritosMap = if (peritosMapRaw.isNotEmpty()) peritosMapRaw else defaultPeritos
                val fotografoMap = if (fotografoMapRaw.isNotEmpty()) fotografoMapRaw else defaultFotografos

                val employeeList = mutableListOf<EmployeeEntity>()

                peritosMap.forEach { (code, name) ->
                    employeeList.add(EmployeeEntity(id = code, name = name, role = "PERITO"))
                }
                fotografoMap.forEach { (code, name) ->
                    employeeList.add(EmployeeEntity(id = 1000 + code, name = name, role = "FOTOGRAFO"))
                }

                // 3. Mapeamento das Colunas Exatas da Tabela do arquivo 01.plantao.trocas.202610.xlsx
                // Peritos: Coluna C (2) = 07h-19h, Coluna D (3) = 19h-07h
                // Fotógrafos: Coluna J (9) = 07h-19h, Coluna K (10) = 19h-07h, Coluna L (11) = Engenharia
                val colPeritoDiurno = 2        // Coluna C
                val colPeritoNoturno = 3       // Coluna D
                val colFotografoDiurno = 9     // Coluna J
                val colFotografoNoturno = 10   // Coluna K
                val colFotografoEngenharia = 11 // Coluna L

                // 4. Processar Escalas Diárias (Linhas dos dias 1 a 31)
                val shiftList = mutableListOf<ShiftEntity>()

                for (row in sheet) {
                    if (row.rowNum < 8) continue // Ignorar cabeçalho inicial

                    // Dia do Mês (1 a 31) na Coluna B (índice 1)
                    val day = parseCellFromRowAsInt(row, 1, formatter) ?: continue

                    if (day in 1..31) {
                        val formattedDate = String.format(Locale.getDefault(), "%02d/%s/%s", day, monthStr, yearStr)

                        // Usar MissingCellPolicy.CREATE_NULL_AS_BLANK para garantir alinhamento físico exato com as colunas
                        val pDiurnoCode = parseCellFromRowAsInt(row, colPeritoDiurno, formatter)
                        val pNoturnoCode = parseCellFromRowAsInt(row, colPeritoNoturno, formatter)

                        val fDiurnoCode = parseCellFromRowAsInt(row, colFotografoDiurno, formatter)
                        val fNoturnoCode = parseCellFromRowAsInt(row, colFotografoNoturno, formatter)
                        val fEngenhariaCode = parseCellFromRowAsInt(row, colFotografoEngenharia, formatter)

                        // --- TURNO 1 (DIURNO - 07h às 19h) ---
                        if (pDiurnoCode != null || fDiurnoCode != null || fEngenhariaCode != null) {
                            shiftList.add(
                                ShiftEntity(
                                    date = formattedDate,
                                    shiftNumber = 1,
                                    plantaoExpertId = pDiurnoCode,
                                    plantaoPhotographerId = fDiurnoCode?.let { 1000 + it },
                                    engenhariaExpertId = null,
                                    engenhariaPhotographerId = fEngenhariaCode?.let { 1000 + it }
                                )
                            )
                        }

                        // --- TURNO 2 (NOTURNO - 19h às 7h) ---
                        if (pNoturnoCode != null || fNoturnoCode != null) {
                            shiftList.add(
                                ShiftEntity(
                                    date = formattedDate,
                                    shiftNumber = 2,
                                    plantaoExpertId = pNoturnoCode,
                                    plantaoPhotographerId = fNoturnoCode?.let { 1000 + it },
                                    engenhariaExpertId = null,
                                    engenhariaPhotographerId = null
                                )
                            )
                        }
                    }
                }

                if (shiftList.isEmpty()) {
                    throw IllegalStateException("Não foi possível extrair a escala. Certifique-se de que a estrutura corresponda à escala da SPTC.")
                }

                // Limpar apenas a escala do mês e ano específicos que estão sendo importados (preservando outros meses no banco!)
                shiftDao.clearShiftsForMonth(monthStr, yearStr)

                if (employeeList.isNotEmpty()) {
                    employeeDao.insertAll(employeeList)
                }
                shiftDao.insertShifts(shiftList)

                "Escala de $monthStr/$yearStr importada com sucesso! (${peritosMap.size} Peritos, ${fotografoMap.size} Fotógrafos e ${shiftList.size} turnos salvos)"
            }
        }
    }

    suspend fun importEngineeringExcelToRoom(uri: Uri): Result<String> = runCatching {
        try {
            context.contentResolver.takePersistableUriPermission(
                uri,
                Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        } catch (_: Exception) {}

        val inputStream: InputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalArgumentException("Não foi possível abrir o arquivo selecionado.")

        val formatter = DataFormatter()

        inputStream.use { stream ->
            val workbook = WorkbookFactory.create(stream)
            workbook.use { wb ->
                if (wb.numberOfSheets == 0) {
                    throw IllegalArgumentException("A planilha de engenharia está vazia.")
                }

                val sheet = wb.getSheetAt(0)

                // 1. Extrair Mês e Ano do cabeçalho
                var monthStr = "10"
                var yearStr = "2026"

                for (rowIndex in 0..12) {
                    val row = sheet.getRow(rowIndex) ?: continue
                    for (c in 0..20) {
                        val text = getCellText(sheet, rowIndex, c, formatter).uppercase()
                        if (text.contains("ENGENHARIA") || text.contains("OUT") || text.contains("202")) {
                            val extracted = parseMonthAndYear(text)
                            if (extracted != null) {
                                monthStr = extracted.first
                                yearStr = extracted.second
                                break
                            }
                        }
                    }
                }

                // 2. Extrair Nomes dos Peritos da Engenharia
                val engPeritosMapRaw = scanSideTable(sheet, formatter, "PERITO")
                val defaultEngPeritos = mapOf(
                    1 to "Feliciano",
                    2 to "Elias C",
                    3 to "Mauro"
                )
                val engPeritosMap = if (engPeritosMapRaw.isNotEmpty()) engPeritosMapRaw else defaultEngPeritos

                val engEmployeeList = mutableListOf<EmployeeEntity>()
                engPeritosMap.forEach { (code, name) ->
                    // Offset +2000 para evitar conflito de IDs com Plantão (+0) e Fotógrafos (+1000)
                    engEmployeeList.add(EmployeeEntity(id = 2000 + code, name = name, role = "PERITO_ENGENHARIA"))
                }
                employeeDao.insertAll(engEmployeeList)

                // 3. Processar linhas dos dias da Engenharia
                // Diurno: Coluna C (2)
                // Noturno: Coluna D (3) ou Coluna E (4) caso a Coluna D esteja oculta na planilha
                var updatedCount = 0

                for (row in sheet) {
                    if (row.rowNum < 8) continue

                    val day = parseCellFromRowAsInt(row, 1, formatter) ?: continue

                    if (day in 1..31) {
                        val formattedDate = String.format(Locale.getDefault(), "%02d/%s/%s", day, monthStr, yearStr)

                        val pEngDiurnoCode = parseCellFromRowAsInt(row, 2, formatter)  // Coluna C
                        val pEngNoturnoCode = parseCellFromRowAsInt(row, 3, formatter) // Coluna D ou E
                            ?: parseCellFromRowAsInt(row, 4, formatter)

                        // Atualizar Turno Diurno (1)
                        if (pEngDiurnoCode != null) {
                            val expertId = 2000 + pEngDiurnoCode
                            val existingShift = shiftDao.getShift(formattedDate, 1)
                            if (existingShift != null) {
                                shiftDao.updateEngenhariaExpert(formattedDate, 1, expertId)
                            } else {
                                shiftDao.insertShifts(
                                    listOf(
                                        ShiftEntity(
                                            date = formattedDate,
                                            shiftNumber = 1,
                                            plantaoExpertId = null,
                                            plantaoPhotographerId = null,
                                            engenhariaExpertId = expertId,
                                            engenhariaPhotographerId = null
                                        )
                                    )
                                )
                            }
                            updatedCount++
                        }

                        // Atualizar Turno Noturno (2) - caso haja algum noturno
                        if (pEngNoturnoCode != null) {
                            val expertId = 2000 + pEngNoturnoCode
                            val existingShift = shiftDao.getShift(formattedDate, 2)
                            if (existingShift != null) {
                                shiftDao.updateEngenhariaExpert(formattedDate, 2, expertId)
                            } else {
                                shiftDao.insertShifts(
                                    listOf(
                                        ShiftEntity(
                                            date = formattedDate,
                                            shiftNumber = 2,
                                            plantaoExpertId = null,
                                            plantaoPhotographerId = null,
                                            engenhariaExpertId = expertId,
                                            engenhariaPhotographerId = null
                                        )
                                    )
                                )
                            }
                            updatedCount++
                        }
                    }
                }

                "Escala de Engenharia de $monthStr/$yearStr importada com sucesso! (${engPeritosMap.size} Peritos de Engenharia vinculados)"
            }
        }
    }

    private fun scanSideTable(sheet: Sheet, formatter: DataFormatter, targetRole: String): Map<Int, String> {
        val map = mutableMapOf<Int, String>()

        for (r in 0..50) {
            for (c in 6..35) {
                val headerText = getCellText(sheet, r, c, formatter)

                // Ignorar a célula "FOTÓGRAFOS / DESENHISTAS" da tabela principal (que fica na coluna J / 9)
                if (c == 9 && headerText.contains("DESENHISTA", ignoreCase = true)) {
                    continue
                }

                val isMatch = if (targetRole == "PERITO") {
                    headerText.equals("Peritos", ignoreCase = true) || cellMatchPerito(headerText)
                } else {
                    headerText.contains("Fotó", ignoreCase = true) || headerText.contains("Fotog", ignoreCase = true) || headerText.contains("Desenhista", ignoreCase = true)
                }

                if (isMatch) {
                    for (subR in (r + 1)..(r + 25)) {
                        val colLeft = getCellText(sheet, subR, c - 1, formatter)
                        val colCenter = getCellText(sheet, subR, c, formatter)
                        val colRight = getCellText(sheet, subR, c + 1, formatter)

                        val idCandidate = colLeft.toIntOrNull() ?: colLeft.toDoubleOrNull()?.toInt()
                            ?: colCenter.toIntOrNull() ?: colCenter.toDoubleOrNull()?.toInt()

                        val nameCandidate = when {
                            colLeft.toIntOrNull() != null || colLeft.toDoubleOrNull() != null -> colCenter
                            colCenter.toIntOrNull() != null || colCenter.toDoubleOrNull() != null -> colRight
                            else -> ""
                        }.trim()

                        val cleanedName = cleanEmployeeName(nameCandidate)

                        if (idCandidate != null && idCandidate in 1..99 && cleanedName.isNotBlank() && !cleanedName.equals("TOTAL", ignoreCase = true)) {
                            map[idCandidate] = cleanedName
                        }
                    }

                    if (map.isNotEmpty()) {
                        return map
                    }
                }
            }
        }
        return map
    }

    private fun cellMatchPerito(text: String): Boolean {
        return text.equals("Perito", ignoreCase = true)
    }

    private fun getCellText(sheet: Sheet, rowIndex: Int, colIndex: Int, formatter: DataFormatter): String {
        val row = sheet.getRow(rowIndex) ?: return ""
        val cell = row.getCell(colIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK) ?: return ""
        return formatter.formatCellValue(cell).trim()
    }

    private fun parseCellFromRowAsInt(row: Row, colIndex: Int, formatter: DataFormatter): Int? {
        val cell = row.getCell(colIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK) ?: return null
        val text = formatter.formatCellValue(cell).trim()
        return text.toIntOrNull() ?: text.toDoubleOrNull()?.toInt()
    }

    private fun parseCellAsInt(cell: Cell?, formatter: DataFormatter): Int? {
        if (cell == null) return null
        val text = formatter.formatCellValue(cell).trim()
        return text.toIntOrNull() ?: text.toDoubleOrNull()?.toInt()
    }

    private fun cleanEmployeeName(name: String): String {
        return name.replace(Regex("^\\d+\\s*"), "").trim()
    }

    private fun parseMonthAndYear(text: String): Pair<String, String>? {
        val months = mapOf(
            "JANEIRO" to "01", "JAN" to "01",
            "FEVEREIRO" to "02", "FEV" to "02",
            "MARÇO" to "03", "MARCO" to "03", "MAR" to "03",
            "ABRIL" to "04", "ABR" to "04",
            "MAIO" to "05", "MAI" to "05",
            "JUNHO" to "06", "JUN" to "06",
            "JULHO" to "07", "JUL" to "07",
            "AGOSTO" to "08", "AGO" to "08",
            "SETEMBRO" to "09", "SET" to "09",
            "OUTUBRO" to "10", "OUT" to "10",
            "NOVEMBRO" to "11", "NOV" to "11",
            "DEZEMBRO" to "12", "DEZ" to "12"
        )
        val upper = text.uppercase()
        var foundMonth: String? = null
        for ((name, num) in months) {
            if (name.length <= 3) {
                val regex = Regex("\\b$name\\b", RegexOption.IGNORE_CASE)
                if (regex.containsMatchIn(upper)) {
                    foundMonth = num
                    break
                }
            } else if (upper.contains(name)) {
                foundMonth = num
                break
            }
        }
        val yearRegex = Regex("20\\d{2}")
        val yearMatch = yearRegex.find(text)
        val foundYear = yearMatch?.value ?: "2026"

        return foundMonth?.let { Pair(it, foundYear) }
    }
}

// ==========================================
// 4. VIEWMODEL E FACTORY
// ==========================================

sealed interface ScheduleUiState {
    data object Idle : ScheduleUiState
    data object Loading : ScheduleUiState
    data class Success(val message: String) : ScheduleUiState
    data class Error(val message: String) : ScheduleUiState
}

class ScheduleViewModel(
    private val shiftDao: ShiftDao,
    private val employeeDao: EmployeeDao,
    private val importer: ScheduleExcelImporter
) : ViewModel() {

    val scheduleList: StateFlow<List<ShiftWithDetails>> = shiftDao.getAllShiftsWithDetails()
        .stateIn(
            scope = viewModelScope,
            started = SharingStarted.WhileSubscribed(5000),
            initialValue = emptyList()
        )

    val allEmployees: StateFlow<List<EmployeeEntity>> = employeeDao.getAllEmployeesFlow()
        .stateIn(
            scope = viewModelScope,
            started = SharingStarted.WhileSubscribed(5000),
            initialValue = emptyList()
        )

    private val _selectedFilterEmployeeIds = MutableStateFlow<Set<Int>>(emptySet())
    val selectedFilterEmployeeIds: StateFlow<Set<Int>> = _selectedFilterEmployeeIds.asStateFlow()

    fun toggleFilterEmployee(employeeId: Int) {
        val current = _selectedFilterEmployeeIds.value.toMutableSet()
        if (current.contains(employeeId)) {
            current.remove(employeeId)
        } else {
            current.add(employeeId)
        }
        _selectedFilterEmployeeIds.value = current
    }

    fun clearFilterEmployees() {
        _selectedFilterEmployeeIds.value = emptySet()
    }

    fun updateShiftStaff(
        shiftId: Long,
        pExpId: Int?,
        pFotId: Int?,
        eExpId: Int?,
        eFotId: Int?
    ) {
        viewModelScope.launch(Dispatchers.IO) {
            shiftDao.updateShiftAssignments(shiftId, pExpId, pFotId, eExpId, eFotId)
        }
    }

    private val _uiState = MutableStateFlow<ScheduleUiState>(ScheduleUiState.Idle)
    val uiState: StateFlow<ScheduleUiState> = _uiState.asStateFlow()

    fun importFile(uri: Uri) {
        viewModelScope.launch(Dispatchers.IO) {
            _uiState.value = ScheduleUiState.Loading
            val result = importer.importExcelToRoom(uri)
            result.onSuccess { msg ->
                _uiState.value = ScheduleUiState.Success(msg)
            }.onFailure { ex ->
                _uiState.value = ScheduleUiState.Error(ex.localizedMessage ?: "Erro desconhecido ao importar planilha.")
            }
        }
    }

    fun importEngenhariaFile(uri: Uri) {
        viewModelScope.launch(Dispatchers.IO) {
            _uiState.value = ScheduleUiState.Loading
            val result = importer.importEngineeringExcelToRoom(uri)
            result.onSuccess { msg ->
                _uiState.value = ScheduleUiState.Success(msg)
            }.onFailure { ex ->
                _uiState.value = ScheduleUiState.Error(ex.localizedMessage ?: "Erro ao importar escala de engenharia.")
            }
        }
    }

    fun dismissState() {
        _uiState.value = ScheduleUiState.Idle
    }
}

class ScheduleViewModelFactory(
    private val shiftDao: ShiftDao,
    private val employeeDao: EmployeeDao,
    private val importer: ScheduleExcelImporter
) : ViewModelProvider.Factory {
    @Suppress("UNCHECKED_CAST")
    override fun <T : ViewModel> create(modelClass: Class<T>): T {
        if (modelClass.isAssignableFrom(ScheduleViewModel::class.java)) {
            return ScheduleViewModel(shiftDao, employeeDao, importer) as T
        }
        throw IllegalArgumentException("Classe ViewModel desconhecida")
    }
}

// ==========================================
// 5. MAIN ACTIVITY E UI COMPOSE
// ==========================================

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        val database = AppDatabase.getDatabase(applicationContext)
        val importer = ScheduleExcelImporter(
            context = applicationContext,
            employeeDao = database.employeeDao(),
            shiftDao = database.shiftDao()
        )
        val viewModelFactory = ScheduleViewModelFactory(database.shiftDao(), database.employeeDao(), importer)

        setContent {
            SPTCEscalaTheme {
                val viewModel: ScheduleViewModel = viewModel(factory = viewModelFactory)
                MainScheduleScreen(viewModel = viewModel)
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun MainScheduleScreen(viewModel: ScheduleViewModel) {
    val shifts by viewModel.scheduleList.collectAsStateWithLifecycle()
    val allEmployees by viewModel.allEmployees.collectAsStateWithLifecycle()
    val selectedFilterIds by viewModel.selectedFilterEmployeeIds.collectAsStateWithLifecycle()
    val uiState by viewModel.uiState.collectAsStateWithLifecycle()
    val snackbarHostState = remember { SnackbarHostState() }
    val coroutineScope = rememberCoroutineScope()

    var isFilterDialogOpen by remember { mutableStateOf(false) }

    if (isFilterDialogOpen) {
        FilterEmployeesDialog(
            allEmployees = allEmployees,
            selectedIds = selectedFilterIds,
            onToggleId = { viewModel.toggleFilterEmployee(it) },
            onClear = { viewModel.clearFilterEmployees() },
            onDismiss = { isFilterDialogOpen = false }
        )
    }

    LaunchedEffect(uiState) {
        when (val state = uiState) {
            is ScheduleUiState.Success -> {
                snackbarHostState.showSnackbar(state.message)
                viewModel.dismissState()
            }
            is ScheduleUiState.Error -> {
                snackbarHostState.showSnackbar(state.message)
                viewModel.dismissState()
            }
            else -> {}
        }
    }

    val plantaoPickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let { viewModel.importFile(it) }
    }

    val engenhariaPickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri?.let { viewModel.importEngenhariaFile(it) }
    }

    val datesList = remember(shifts) {
        shifts.map { it.shift.date }.distinct().sortedBy { dateStr ->
            val parts = dateStr.split("/")
            if (parts.size == 3) {
                "${parts[2]}${parts[1]}${parts[0]}"
            } else {
                dateStr
            }
        }
    }

    val todayStr = remember {
        val sdf = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault())
        sdf.format(Date())
    }

    val initialPageIndex = remember(datesList) {
        if (datesList.isEmpty()) return@remember 0

        val exactTodayIdx = datesList.indexOf(todayStr)
        if (exactTodayIdx >= 0) return@remember exactTodayIdx

        val todayParts = todayStr.split("/")
        if (todayParts.size == 3) {
            val dayPart = todayParts[0]
            val monthPart = todayParts[1]

            val sameDayMonthIdx = datesList.indexOfFirst { it.startsWith("$dayPart/$monthPart/") }
            if (sameDayMonthIdx >= 0) return@remember sameDayMonthIdx

            val sameMonthIdx = datesList.indexOfFirst { "/$monthPart/" in it }
            if (sameMonthIdx >= 0) return@remember sameMonthIdx
        }

        0
    }

    val pagerState = rememberPagerState(
        initialPage = initialPageIndex,
        pageCount = { datesList.size }
    )

    var hasInitializedPage by remember { mutableStateOf(false) }

    LaunchedEffect(datesList, initialPageIndex) {
        if (datesList.isNotEmpty() && !hasInitializedPage) {
            hasInitializedPage = true
            if (initialPageIndex in datesList.indices) {
                pagerState.scrollToPage(initialPageIndex)
            }
        }
    }

    Scaffold(
        snackbarHost = { SnackbarHost(snackbarHostState) },
        topBar = {
            TopAppBar(
                title = { Text("SPTC-Escalas") },
                actions = {
                    IconButton(onClick = { isFilterDialogOpen = true }) {
                        BadgedBox(
                            badge = {
                                if (selectedFilterIds.isNotEmpty()) {
                                    Badge { Text("${selectedFilterIds.size}") }
                                }
                            }
                        ) {
                            Icon(
                                imageVector = Icons.Default.FilterList,
                                contentDescription = "Filtrar Servidores"
                            )
                        }
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.primaryContainer,
                    titleContentColor = MaterialTheme.colorScheme.onPrimaryContainer
                )
            )
        },
        bottomBar = {
            Surface(
                tonalElevation = 3.dp,
                color = MaterialTheme.colorScheme.surface
            ) {
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 16.dp, vertical = 8.dp),
                    horizontalArrangement = Arrangement.SpaceEvenly,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Button(
                        onClick = { plantaoPickerLauncher.launch(arrayOf("*/*")) },
                        colors = ButtonDefaults.buttonColors(
                            containerColor = MaterialTheme.colorScheme.primaryContainer,
                            contentColor = MaterialTheme.colorScheme.onPrimaryContainer
                        )
                    ) {
                        Icon(Icons.Default.FileOpen, contentDescription = null, modifier = Modifier.size(18.dp))
                        Spacer(modifier = Modifier.width(6.dp))
                        Text("Plantão")
                    }

                    Button(
                        onClick = { engenhariaPickerLauncher.launch(arrayOf("*/*")) },
                        colors = ButtonDefaults.buttonColors(
                            containerColor = MaterialTheme.colorScheme.secondaryContainer,
                            contentColor = MaterialTheme.colorScheme.onSecondaryContainer
                        )
                    ) {
                        Icon(Icons.Default.Engineering, contentDescription = null, modifier = Modifier.size(18.dp))
                        Spacer(modifier = Modifier.width(6.dp))
                        Text("Engenharia")
                    }
                }
            }
        }
    ) { innerPadding ->
        Box(
            modifier = Modifier
                .fillMaxSize()
                .padding(innerPadding)
        ) {
            if (shifts.isEmpty() && (uiState !is ScheduleUiState.Loading)) {
                Box(
                    modifier = Modifier.fillMaxSize(),
                    contentAlignment = Alignment.Center
                ) {
                    Text(
                        text = "Nenhuma escala carregada.\nImporte a escala de Plantão e a de Engenharia.",
                        style = MaterialTheme.typography.bodyLarge,
                        modifier = Modifier.padding(16.dp)
                    )
                }
            } else {
                Column(modifier = Modifier.fillMaxSize()) {
                    val selectedDate = if (datesList.isNotEmpty() && pagerState.currentPage < datesList.size) {
                        datesList[pagerState.currentPage]
                    } else null

                    MonthCalendar(
                        datesList = datesList,
                        selectedDate = selectedDate,
                        shifts = shifts,
                        selectedFilterIds = selectedFilterIds,
                        onDateSelected = { targetIndex ->
                            coroutineScope.launch {
                                pagerState.animateScrollToPage(targetIndex)
                            }
                        }
                    )

                    HorizontalPager(
                        state = pagerState,
                        modifier = Modifier
                            .fillMaxWidth()
                            .weight(1f),
                        contentPadding = PaddingValues(horizontal = 16.dp, vertical = 8.dp),
                        pageSpacing = 16.dp
                    ) { pageIndex ->
                        val currentDate = datesList[pageIndex]
                        val dailyShifts = shifts.filter { it.shift.date == currentDate }

                        LazyColumn(
                            modifier = Modifier.fillMaxSize(),
                            verticalArrangement = Arrangement.spacedBy(12.dp)
                        ) {
                            items(
                                items = dailyShifts,
                                key = { shiftDetails -> shiftDetails.shift.id }
                            ) { item ->
                                ShiftCard(
                                    data = item,
                                    allEmployees = allEmployees,
                                    selectedFilterIds = selectedFilterIds,
                                    onUpdateStaff = { pExpId, pFotId, eExpId, eFotId ->
                                        viewModel.updateShiftStaff(item.shift.id, pExpId, pFotId, eExpId, eFotId)
                                    }
                                )
                            }
                        }
                    }
                }
            }

            if (uiState is ScheduleUiState.Loading) {
                CircularProgressIndicator(modifier = Modifier.align(Alignment.Center))
            }
        }
    }
}

@Composable
fun MonthCalendar(
    datesList: List<String>,
    selectedDate: String?,
    shifts: List<ShiftWithDetails>,
    selectedFilterIds: Set<Int>,
    onDateSelected: (Int) -> Unit
) {
    if (datesList.isEmpty()) return

    val sampleDate = datesList.firstOrNull() ?: "01/10/2026"
    val parts = sampleDate.split("/")
    val defaultMonthNum = parts.getOrNull(1)?.toIntOrNull() ?: 10
    val defaultYearNum = parts.getOrNull(2)?.toIntOrNull() ?: 2026

    var displayedMonth by remember(sampleDate) { mutableIntStateOf(defaultMonthNum) }
    var displayedYear by remember(sampleDate) { mutableIntStateOf(defaultYearNum) }

    LaunchedEffect(selectedDate) {
        selectedDate?.split("/")?.let { dateParts ->
            val m = dateParts.getOrNull(1)?.toIntOrNull()
            val y = dateParts.getOrNull(2)?.toIntOrNull()
            if (m != null && y != null) {
                displayedMonth = m
                displayedYear = y
            }
        }
    }

    val monthNames = arrayOf(
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    )
    val monthTitle = "${monthNames.getOrElse(displayedMonth - 1) { "Outubro" }} - $displayedYear"

    fun switchMonth(monthDelta: Int) {
        var newM = displayedMonth + monthDelta
        var newY = displayedYear
        if (newM > 12) {
            newM = 1
            newY += 1
        } else if (newM < 1) {
            newM = 12
            newY -= 1
        }
        displayedMonth = newM
        displayedYear = newY

        val monthPattern = String.format(Locale.getDefault(), "/%02d/%04d", newM, newY)
        val targetIndex = datesList.indexOfFirst { it.endsWith(monthPattern) }
        if (targetIndex >= 0) {
            onDateSelected(targetIndex)
        }
    }

    val daysWithData = remember(datesList, displayedMonth, displayedYear) {
        datesList.filter {
            val p = it.split("/")
            p.getOrNull(1)?.toIntOrNull() == displayedMonth && p.getOrNull(2)?.toIntOrNull() == displayedYear
        }.mapNotNull { it.split("/").getOrNull(0)?.toIntOrNull() }.toSet()
    }

    val selectedDayNum = remember(selectedDate, displayedMonth, displayedYear) {
        selectedDate?.split("/")?.let { p ->
            if (p.getOrNull(1)?.toIntOrNull() == displayedMonth && p.getOrNull(2)?.toIntOrNull() == displayedYear) {
                p.getOrNull(0)?.toIntOrNull()
            } else null
        }
    }

    var dragSum by remember { mutableFloatStateOf(0f) }

    Card(
        modifier = Modifier
            .fillMaxWidth()
            .padding(horizontal = 16.dp, vertical = 6.dp)
            .pointerInput(Unit) {
                detectHorizontalDragGestures(
                    onDragStart = { dragSum = 0f },
                    onDragEnd = {
                        if (dragSum < -50f) {
                            switchMonth(+1)
                        } else if (dragSum > 50f) {
                            switchMonth(-1)
                        }
                    },
                    onHorizontalDrag = { change, dragAmount ->
                        change.consume()
                        dragSum += dragAmount
                    }
                )
            },
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.45f))
    ) {
        Column(
            modifier = Modifier.padding(horizontal = 12.dp, vertical = 10.dp),
            horizontalAlignment = Alignment.CenterHorizontally
        ) {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                IconButton(
                    onClick = { switchMonth(-1) },
                    modifier = Modifier.size(32.dp)
                ) {
                    Icon(Icons.Default.ChevronLeft, contentDescription = "Mês Anterior", modifier = Modifier.size(20.dp))
                }

                Text(
                    text = "📅 $monthTitle",
                    style = MaterialTheme.typography.titleMedium,
                    color = MaterialTheme.colorScheme.primary
                )

                IconButton(
                    onClick = { switchMonth(+1) },
                    modifier = Modifier.size(32.dp)
                ) {
                    Icon(Icons.Default.ChevronRight, contentDescription = "Próximo Mês", modifier = Modifier.size(20.dp))
                }
            }

            Spacer(modifier = Modifier.height(4.dp))

            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceAround
            ) {
                listOf("D", "S", "T", "Q", "Q", "S", "S").forEach { dayHeader ->
                    Text(
                        text = dayHeader,
                        style = MaterialTheme.typography.labelMedium,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }

            Spacer(modifier = Modifier.height(4.dp))

            val totalDays = 31
            val rowsCount = (totalDays + 6) / 7

            Column(verticalArrangement = Arrangement.spacedBy(3.dp)) {
                for (r in 0 until rowsCount) {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.SpaceAround
                    ) {
                        for (c in 0..6) {
                            val dayNum = r * 7 + c + 1
                            if (dayNum <= totalDays) {
                                val isSelected = dayNum == selectedDayNum
                                val hasShift = daysWithData.contains(dayNum)

                                val targetDateStr = String.format(Locale.getDefault(), "%02d/%02d/%04d", dayNum, displayedMonth, displayedYear)

                                val dayShifts = remember(shifts, targetDateStr) {
                                    shifts.filter { it.shift.date == targetDateStr }
                                }

                                var hasPlantaoPerito = false
                                var hasPlantaoFotografo = false
                                var hasEngenhariaPerito = false
                                var hasEngenhariaFotografo = false

                                if (selectedFilterIds.isNotEmpty()) {
                                    dayShifts.forEach { item ->
                                        if (item.shift.plantaoExpertId in selectedFilterIds) hasPlantaoPerito = true
                                        if (item.shift.plantaoPhotographerId in selectedFilterIds) hasPlantaoFotografo = true
                                        if (item.shift.engenhariaExpertId in selectedFilterIds) hasEngenhariaPerito = true
                                        if (item.shift.engenhariaPhotographerId in selectedFilterIds) hasEngenhariaFotografo = true
                                    }
                                }

                                val isFilterMatched = hasPlantaoPerito || hasPlantaoFotografo || hasEngenhariaPerito || hasEngenhariaFotografo

                                Surface(
                                    onClick = {
                                        val targetIdx = datesList.indexOf(targetDateStr)
                                        if (targetIdx >= 0) {
                                            onDateSelected(targetIdx)
                                        }
                                    },
                                    shape = CircleShape,
                                    color = when {
                                        isSelected -> MaterialTheme.colorScheme.primary
                                        isFilterMatched -> MaterialTheme.colorScheme.secondaryContainer
                                        hasShift -> MaterialTheme.colorScheme.surface
                                        else -> MaterialTheme.colorScheme.surface
                                    },
                                    contentColor = when {
                                        isSelected -> MaterialTheme.colorScheme.onPrimary
                                        isFilterMatched -> MaterialTheme.colorScheme.onSecondaryContainer
                                        hasShift -> MaterialTheme.colorScheme.onSurface
                                        else -> MaterialTheme.colorScheme.onSurface
                                    },
                                    border = if (isFilterMatched && !isSelected) {
                                        BorderStroke(1.5.dp, MaterialTheme.colorScheme.primary)
                                    } else null,
                                    modifier = Modifier.size(34.dp)
                                ) {
                                    Column(
                                        horizontalAlignment = Alignment.CenterHorizontally,
                                        verticalArrangement = Arrangement.Center
                                    ) {
                                        Text(
                                            text = "$dayNum",
                                            style = MaterialTheme.typography.bodySmall
                                        )

                                        if (isFilterMatched) {
                                            Row(
                                                horizontalArrangement = Arrangement.spacedBy(1.dp),
                                                verticalAlignment = Alignment.CenterVertically
                                            ) {
                                                if (hasPlantaoPerito) {
                                                    Box(
                                                        modifier = Modifier
                                                            .size(4.dp)
                                                            .clip(CircleShape)
                                                            .background(Color(0xFF2E7D32))
                                                    )
                                                }
                                                if (hasPlantaoFotografo) {
                                                    Box(
                                                        modifier = Modifier
                                                            .size(4.dp)
                                                            .clip(CircleShape)
                                                            .background(Color(0xFF1565C0))
                                                    )
                                                }
                                                if (hasEngenhariaPerito || hasEngenhariaFotografo) {
                                                    Box(
                                                        modifier = Modifier
                                                            .size(4.dp)
                                                            .clip(CircleShape)
                                                            .background(Color(0xFFE65100))
                                                    )
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                Spacer(modifier = Modifier.size(34.dp))
                            }
                        }
                    }
                }
            }
        }
    }
}

@Composable
fun HighlightedStaffText(
    roleLabel: String,
    employee: EmployeeEntity?,
    selectedFilterIds: Set<Int>
) {
    val isSelected = employee != null && selectedFilterIds.contains(employee.id)

    Row(
        verticalAlignment = Alignment.CenterVertically,
        horizontalArrangement = Arrangement.spacedBy(4.dp)
    ) {
        Text(
            text = "• $roleLabel: ",
            style = MaterialTheme.typography.bodyMedium
        )
        if (employee != null) {
            if (isSelected) {
                Surface(
                    color = MaterialTheme.colorScheme.primaryContainer,
                    shape = MaterialTheme.shapes.extraSmall
                ) {
                    Text(
                        text = " ${employee.name} ⭐ ",
                        style = MaterialTheme.typography.bodyMedium.copy(fontWeight = FontWeight.Bold),
                        color = MaterialTheme.colorScheme.onPrimaryContainer
                    )
                }
            } else {
                Text(
                    text = employee.name,
                    style = MaterialTheme.typography.bodyMedium
                )
            }
        } else {
            Text(
                text = "Não escalado",
                style = MaterialTheme.typography.bodyMedium,
                color = MaterialTheme.colorScheme.onSurfaceVariant
            )
        }
    }
}

@Composable
fun ShiftCard(
    data: ShiftWithDetails,
    allEmployees: List<EmployeeEntity>,
    selectedFilterIds: Set<Int>,
    onUpdateStaff: (pExpId: Int?, pFotId: Int?, eExpId: Int?, eFotId: Int?) -> Unit
) {
    var isEditing by remember { mutableStateOf(false) }

    val isShiftOne = data.shift.shiftNumber == 1
    val shiftTitle = if (isShiftOne) "1º Turno (Diurno - 07h às 19h)" else "2º Turno (Noturno - 19h às 7h)"
    val shiftIcon = if (isShiftOne) Icons.Default.WbSunny else Icons.Default.Nightlight

    if (isEditing) {
        EditShiftDialog(
            data = data,
            allEmployees = allEmployees,
            onDismiss = { isEditing = false },
            onSave = { pExpId, pFotId, eExpId, eFotId ->
                onUpdateStaff(pExpId, pFotId, eExpId, eFotId)
                isEditing = false
            }
        )
    }

    Card(
        modifier = Modifier.fillMaxWidth(),
        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp)
    ) {
        Column(modifier = Modifier.padding(horizontal = 16.dp, vertical = 12.dp)) {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    horizontalArrangement = Arrangement.spacedBy(8.dp)
                ) {
                    Icon(
                        imageVector = shiftIcon,
                        contentDescription = null,
                        tint = MaterialTheme.colorScheme.primary,
                        modifier = Modifier.size(20.dp)
                    )
                    Text(
                        text = shiftTitle,
                        style = MaterialTheme.typography.titleMedium,
                        color = MaterialTheme.colorScheme.primary
                    )
                }

                IconButton(
                    onClick = { isEditing = true },
                    modifier = Modifier.size(28.dp)
                ) {
                    Icon(
                        imageVector = Icons.Default.Edit,
                        contentDescription = "Editar Troca / Plantão",
                        tint = MaterialTheme.colorScheme.primary,
                        modifier = Modifier.size(18.dp)
                    )
                }
            }

            Spacer(modifier = Modifier.height(6.dp))
            HorizontalDivider()
            Spacer(modifier = Modifier.height(6.dp))

            // Equipe de Plantão
            Text(
                text = "🚨 Plantão",
                style = MaterialTheme.typography.labelLarge,
                color = MaterialTheme.colorScheme.secondary
            )
            HighlightedStaffText("Perito", data.plantaoExpert, selectedFilterIds)
            HighlightedStaffText("Fotógrafo", data.plantaoPhotographer, selectedFilterIds)

            // Equipe de Engenharia
            val hasEngenharia = data.engenhariaExpert != null || data.engenhariaPhotographer != null || isShiftOne
            if (hasEngenharia) {
                Spacer(modifier = Modifier.height(8.dp))
                Text(
                    text = "⚙️ Engenharia",
                    style = MaterialTheme.typography.labelLarge,
                    color = MaterialTheme.colorScheme.tertiary
                )
                HighlightedStaffText("Perito", data.engenhariaExpert, selectedFilterIds)
                if (isShiftOne || data.engenhariaPhotographer != null) {
                    HighlightedStaffText("Fotógrafo", data.engenhariaPhotographer, selectedFilterIds)
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class, ExperimentalLayoutApi::class)
@Composable
fun FilterEmployeesDialog(
    allEmployees: List<EmployeeEntity>,
    selectedIds: Set<Int>,
    onToggleId: (Int) -> Unit,
    onClear: () -> Unit,
    onDismiss: () -> Unit
) {
    val peritos = remember(allEmployees) {
        allEmployees.filter { it.role == "PERITO" || it.role == "PERITO_ENGENHARIA" }
    }
    val fotografos = remember(allEmployees) {
        allEmployees.filter { it.role == "FOTOGRAFO" }
    }

    AlertDialog(
        onDismissRequest = onDismiss,
        title = {
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.SpaceBetween,
                verticalAlignment = Alignment.CenterVertically
            ) {
                Text("🔍 Filtrar Pessoas")
                if (selectedIds.isNotEmpty()) {
                    TextButton(onClick = onClear) {
                        Text("Limpar (${selectedIds.size})", style = MaterialTheme.typography.labelMedium)
                    }
                }
            }
        },
        text = {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .verticalScroll(rememberScrollState()),
                verticalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                Text("📋 Peritos:", style = MaterialTheme.typography.labelLarge, color = MaterialTheme.colorScheme.secondary)
                FlowRow(
                    horizontalArrangement = Arrangement.spacedBy(6.dp),
                    verticalArrangement = Arrangement.spacedBy(6.dp)
                ) {
                    peritos.forEach { emp ->
                        val isSelected = selectedIds.contains(emp.id)
                        FilterChip(
                            selected = isSelected,
                            onClick = { onToggleId(emp.id) },
                            label = { Text(emp.name) },
                            leadingIcon = if (isSelected) {
                                { Icon(Icons.Default.Check, contentDescription = null, modifier = Modifier.size(16.dp)) }
                            } else null
                        )
                    }
                }

                Spacer(modifier = Modifier.height(4.dp))
                Text("📷 Fotógrafos:", style = MaterialTheme.typography.labelLarge, color = MaterialTheme.colorScheme.tertiary)
                FlowRow(
                    horizontalArrangement = Arrangement.spacedBy(6.dp),
                    verticalArrangement = Arrangement.spacedBy(6.dp)
                ) {
                    fotografos.forEach { emp ->
                        val isSelected = selectedIds.contains(emp.id)
                        FilterChip(
                            selected = isSelected,
                            onClick = { onToggleId(emp.id) },
                            label = { Text(emp.name) },
                            leadingIcon = if (isSelected) {
                                { Icon(Icons.Default.Check, contentDescription = null, modifier = Modifier.size(16.dp)) }
                            } else null
                        )
                    }
                }
            }
        },
        confirmButton = {
            Button(onClick = onDismiss) {
                Text("Concluir")
            }
        }
    )
}

@Composable
fun EditShiftDialog(
    data: ShiftWithDetails,
    allEmployees: List<EmployeeEntity>,
    onDismiss: () -> Unit,
    onSave: (pExpId: Int?, pFotId: Int?, eExpId: Int?, eFotId: Int?) -> Unit
) {
    val peritosList = remember(allEmployees) {
        allEmployees.filter { it.role == "PERITO" || it.role == "PERITO_ENGENHARIA" }
    }
    val fotografoList = remember(allEmployees) {
        allEmployees.filter { it.role == "FOTOGRAFO" }
    }

    var selectedPExpId by remember { mutableStateOf(data.shift.plantaoExpertId) }
    var selectedPFotId by remember { mutableStateOf(data.shift.plantaoPhotographerId) }
    var selectedEExpId by remember { mutableStateOf(data.shift.engenhariaExpertId) }
    var selectedEFotId by remember { mutableStateOf(data.shift.engenhariaPhotographerId) }

    val isShiftOne = data.shift.shiftNumber == 1

    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text("Editar Plantão - ${data.shift.date} (${if (isShiftOne) "Diurno" else "Noturno"})") },
        text = {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .verticalScroll(rememberScrollState()),
                verticalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                Text("🚨 Plantão:", style = MaterialTheme.typography.labelLarge, color = MaterialTheme.colorScheme.secondary)

                EmployeeDropdown(
                    label = "Perito do Plantão",
                    selectedId = selectedPExpId,
                    options = peritosList,
                    onSelected = { selectedPExpId = it }
                )

                EmployeeDropdown(
                    label = "Fotógrafo do Plantão",
                    selectedId = selectedPFotId,
                    options = fotografoList,
                    onSelected = { selectedPFotId = it }
                )

                Spacer(modifier = Modifier.height(8.dp))
                Text("⚙️ Engenharia:", style = MaterialTheme.typography.labelLarge, color = MaterialTheme.colorScheme.tertiary)

                EmployeeDropdown(
                    label = "Perito da Engenharia",
                    selectedId = selectedEExpId,
                    options = peritosList,
                    onSelected = { selectedEExpId = it }
                )

                if (isShiftOne) {
                    EmployeeDropdown(
                        label = "Fotógrafo da Engenharia",
                        selectedId = selectedEFotId,
                        options = fotografoList,
                        onSelected = { selectedEFotId = it }
                    )
                }
            }
        },
        confirmButton = {
            TextButton(onClick = {
                onSave(selectedPExpId, selectedPFotId, selectedEExpId, selectedEFotId)
            }) {
                Text("Salvar")
            }
        },
        dismissButton = {
            TextButton(onClick = onDismiss) {
                Text("Cancelar")
            }
        }
    )
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun EmployeeDropdown(
    label: String,
    selectedId: Int?,
    options: List<EmployeeEntity>,
    onSelected: (Int?) -> Unit
) {
    var expanded by remember { mutableStateOf(false) }
    val selectedName = options.find { it.id == selectedId }?.name ?: "Não escalado"

    ExposedDropdownMenuBox(
        expanded = expanded,
        onExpandedChange = { expanded = !expanded }
    ) {
        OutlinedTextField(
            value = selectedName,
            onValueChange = {},
            readOnly = true,
            label = { Text(label) },
            trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = expanded) },
            modifier = Modifier
                .fillMaxWidth()
                .menuAnchor()
        )
        ExposedDropdownMenu(
            expanded = expanded,
            onDismissRequest = { expanded = false }
        ) {
            DropdownMenuItem(
                text = { Text("Não escalado") },
                onClick = {
                    onSelected(null)
                    expanded = false
                }
            )
            options.forEach { employee ->
                DropdownMenuItem(
                    text = { Text(employee.name) },
                    onClick = {
                        onSelected(employee.id)
                        expanded = false
                    }
                )
            }
        }
    }
}
