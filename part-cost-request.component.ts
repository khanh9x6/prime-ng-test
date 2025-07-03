import { DateHelperService } from './../../services/date-helper.service'
import { Router } from '@angular/router'
import jsPDF from 'jspdf'
import html2canvas from 'html2canvas'
import * as XLSX from 'xlsx'
import { parse, isValid, isBefore, isAfter, isEqual, startOfDay } from 'date-fns'
import {
  ActionData,
  ActionRequestEnum,
  FileParameter,
  MasterChangePriceReason,
  MasterCrCuReason,
  MasterCurrency,
  MasterEmployee,
  MasterIncoterm,
  MasterOriginal,
  MasterRequestAction,
  MasterRequestFilter,
  MasterRequestStatus,
  MasterRouting,
  MasterSupplier,
  MasterTypeFormRequestPartCost,
  MasterTypeFormRequestPartCostEnum,
  MasterTypeRequestPartCost,
  MasterTypeRequestPartCostEnum,
  MasterUnit,
  PartCostRequestData,
  PartCostRequestDataResponseData,
  PartCostService,
  RequestBase,
  RequestBaseSupportDocument,
  RequestPartCost,
  RequestPartCostDetail,
  ResponseData,
  StatusRequestEnum
} from './../../services/api.service.generated'
import {
  Component,
  ElementRef,
  HostListener,
  Input,
  OnDestroy,
  OnInit,
  Renderer2,
  ViewChild
} from '@angular/core'
import { ConfirmationService, MessageService } from 'primeng/api'
import { MasterService } from 'src/app/services/api.service.generated'
import { HelperService } from 'src/app/services/helper.service'
import { AuthClientService } from 'src/app/services/auth.client.service'
import { DatePipe } from '@angular/common'
import { FileUpload } from 'primeng/fileupload'
import { debounceTime, Subject, Subscription } from 'rxjs'
import Swal from 'sweetalert2'

@Component({
    selector: 'app-part-cost-request',
    templateUrl: './part-cost-request.component.html',
    styleUrl: './part-cost-request.component.scss',
    providers: [MessageService, DatePipe],
    standalone: false
})
export class PartCostRequestComponent implements OnInit, OnDestroy {
  lengthPartCodeRequired: number = 50
  isLoading: boolean = false
  isEditMode: boolean = false
  isRequiredRemark: boolean = false
  arrayEmployeeCodeCanRfi: string[] = []
  	  showPrintPreview: boolean = false;
  printMenuItems: any[] = [];
  @Input() isView: boolean = false
  @Input() requestId: string = ''
  @ViewChild('excelFileUpload') fileUpload: FileUpload
  clonedPartCosts: { [s: string]: RequestPartCostDetail } = {}
  uploadFilesConvert: FileParameter[] = []
  // For validation
  submitted = false
  errors: { [rowId: string]: { [fieldName: string]: boolean } } = {}
  errorsGeneral: { [fieldName: string]: boolean } = {}
  isChangeTypeRequestPartCost: boolean = false
  masterEmployees: MasterEmployee[] = []
  masterEmployeesMap: Map<string, MasterEmployee> = new Map<string, MasterEmployee>()
  masterSuppliers: MasterSupplier[] = []
  masterRoutings: MasterRouting[] = []
  masterSuppliersFilter: MasterSupplier[] = []
  masterSuppliersMap: Map<string, MasterSupplier> = new Map<string, MasterSupplier>()
  partCostRequestData: PartCostRequestData = new PartCostRequestData()
  masterChangePriceReasons: MasterChangePriceReason[] = []
  masterChangePriceReasonsFilter: MasterChangePriceReason[] = []
  masterChangePriceReasonsMap: Map<number, MasterChangePriceReason> = new Map<
    number,
    MasterChangePriceReason
  >()
  masterTypeRequestPartCosts: MasterTypeRequestPartCost[] = []
  masterTypeRequestPartCostsFilter: MasterTypeRequestPartCost[] = []
  masterTypeRequestPartCostsMap: Map<number, MasterTypeRequestPartCost> = new Map<
    number,
    MasterTypeRequestPartCost
  >()
  uploadedFiles: any[] = []
  uploadedQuotationFiles: any[] = []
  masterTypeFormRequestPartCosts: MasterTypeFormRequestPartCost[] = []
  masterTypeFormRequestPartCostsFilter: MasterTypeFormRequestPartCost[] = []
  masterTypeFormRequestPartCostsMap: Map<number, MasterTypeFormRequestPartCost> = new Map<
    number,
    MasterTypeFormRequestPartCost
  >()
  masterCurrencies: MasterCurrency[] = []
  currencyPayment: MasterCurrency | null = null
  currencyPart: MasterCurrency | null = null
  masterCurrenciesFilter: MasterCurrency[] = []
  masterCurrenciesMap: Map<number, MasterCurrency> = new Map<number, MasterCurrency>()

  masterIncoterms: MasterIncoterm[] = []
  masterIncotermsFilter: MasterIncoterm[] = []
  masterIncotermsMap: Map<number, MasterIncoterm> = new Map<number, MasterIncoterm>()

  masterOriginals: MasterOriginal[] = []
  masterOriginalsFilter: MasterOriginal[] = []
  masterOriginalsMap: Map<number, MasterOriginal> = new Map<number, MasterOriginal>()

  masterCrCuReasons: MasterCrCuReason[] = []
  masterCrCuReasonsFilter: MasterCrCuReason[] = []
  masterCrCuReasonsMap: Map<number, MasterCrCuReason> = new Map<number, MasterCrCuReason>()

  masterUnits: MasterUnit[] = []
  masterUnitsFilter: MasterUnit[] = []
  masterUnitsMap: Map<number, MasterUnit> = new Map<number, MasterUnit>()

  masterRequestActions: MasterRequestAction[] = []
  masterRequestActionsFilter: MasterRequestAction[] = []
  masterRequestActionsMap: Map<number, MasterRequestAction> = new Map<number, MasterRequestAction>()

  masterRequestFilters: MasterRequestFilter[] = []
  masterRequestFiltersFilter: MasterRequestFilter[] = []
  masterRequestFiltersMap: Map<number, MasterRequestFilter> = new Map<number, MasterRequestFilter>()

  masterRequestStatuses: MasterRequestStatus[] = []
  masterRequestStatusesFilter: MasterRequestStatus[] = []
  masterRequestStatusesMap: Map<number, MasterRequestStatus> = new Map<
    number,
    MasterRequestStatus
  >()
  actionData: ActionData = new ActionData()
  comment: string = ''
  visible: boolean = false
  isFullScreen = false
  selectedItems: RequestPartCostDetail[] = []
  isViewMode: boolean = false
  firstSubmitDate: Date = new Date()

  @ViewChild('inputRef') inputRef!: ElementRef<HTMLInputElement>
  @ViewChild('sizerRef') sizerRef!: ElementRef<HTMLElement>
  ngOnDestroy() {
    // Khôi phục scroll khi component bị destroy
    this.renderer.removeStyle(document.body, 'overflow')

    // Cleanup all subscriptions
    this.subscriptions.forEach((sub) => sub.unsubscribe())
  }
  toggleFullScreen() {
    this.isFullScreen = !this.isFullScreen
    if (this.isFullScreen) {
      this.renderer.setStyle(document.body, 'overflow', 'hidden') // disable scroll
    } else {
      this.renderer.removeStyle(document.body, 'overflow') // enable scroll
    }
  }
  showDialog() {
    this.visible = true
  }
  constructor(
    private masterService: MasterService,
    private messageService: MessageService,
    public helperService: HelperService,
    private partCostService: PartCostService,
    public authClientService: AuthClientService,
    private router: Router,
    private renderer: Renderer2,
    private dateHelperService: DateHelperService,
    private datePipe: DatePipe,
    private confirmationService: ConfirmationService
    // Inject DatePipe
  ) {
    // console.log(this.partCostRequestData)
  }
  dropAllItem() {
    this.confirmationService.confirm({
      message: 'Are you sure you want to delete all items? This action cannot be undone.',
      header: 'Confirm Delete',
      icon: 'pi pi-exclamation-triangle',
      accept: () => {
        this.partCostRequestData.requestPartCostDetails = []
        this.messageService.add({
          severity: 'success',
          summary: 'Success',
          detail: 'Delete Selected Items'
        })
      },
      reject: () => {
        // User cancelled, do nothing
      }
    })
  }

  ngOnInit() {
    this.getMasterCrCuReason()
    this.getMasterChangePriceReason()
    this.getMasterCurrency()
    this.getMasterIncoterm()
    this.getMasterOriginal()
    this.getMasterRequestAction()
    this.getMasterRequestFilter()
    this.getMasterRequestStatus()
    this.getMasterSupplier()
    this.getMasterTypeRequestPartCost()
    this.getMasterTypeFormRequestPartCost()
    this.masterUnit()
    // this.masterRouting()
    this.initPartCostRequestData()
    this.masterEmployee()
    this.masterRoutingForEmployeeCode(this.authClientService.getEmployeeCodeAuthed())
    if (this.isView) {
      // this.initPartCostRequestData()
      this.actionData.comment = ''
      this.actionData.requestId = this.requestId
      this.actionData.masterRoutingLevel = this.partCostRequestData.requestBase.masterRoutingLevel
      this.isLoading = true
      setTimeout(() => this.getData(), 1500)
    }
    // Setup debounced validation for part code changes
    const validationSubscription = this.validationSubject
      .pipe(debounceTime(300))
      .subscribe(({ partCost, field }) => {
        this.performValidation(partCost, field)
      })
    this.subscriptions.push(validationSubscription)
  }

  private validationSubject = new Subject<{ partCost: RequestPartCostDetail; field: string }>()
  private costChangeSubject = new Subject<RequestPartCostDetail>()
  private subscriptions: Subscription[] = []

  getData() {
    console.log('getData')
    this.partCostService.getPartCostRequest(this.requestId).subscribe({
      next: (data: PartCostRequestDataResponseData) => {
        this.isLoading = false
        if (data.isError == false) {
          this.partCostRequestData = data.result
          this.partCostRequestData.masterRoutings = []
          this.partCostRequestData.requestApprovalSummaries = data.result.requestApprovalSummaries
          this.partCostRequestData.requestBase = data.result.requestBase
          this.partCostRequestData.requestBase.requestNo = data.result.requestBase.requestNo
          this.partCostRequestData.requestBaseSupportDocuments =
            data.result.requestBaseSupportDocuments
          this.partCostRequestData.requestHistoryRoutings = data.result.requestHistoryRoutings
          this.partCostRequestData.requestPartCostDetails = data.result.requestPartCostDetails
          this.partCostRequestData.masterChangePriceReasonLogs =
            data.result.masterChangePriceReasonLogs
          this.partCostRequestData.requestPartCost = data.result.requestPartCost
          this.calculateFirstSubmitDate()
          this.currencyPart = this.masterCurrenciesMap.has(
            this.partCostRequestData.requestPartCost.partMasterCurrencyId
          )
            ? this.masterCurrenciesMap.get(
                this.partCostRequestData.requestPartCost.partMasterCurrencyId
              )
            : null
          this.currencyPayment = this.masterCurrenciesMap.has(
            this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
          )
            ? this.masterCurrenciesMap.get(
                this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
              )
            : null
          for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
            for (
              let j = 0;
              j <
              this.partCostRequestData.requestPartCostDetails[i].requestPartCostDetailCrCuReasons
                .length;
              j++
            ) {
              let idNo =
                this.partCostRequestData.requestPartCostDetails[i].requestPartCostDetailCrCuReasons[
                  j
                ].masterCrCuReasonIdNo
              if (this.masterCrCuReasonsMap.has(idNo)) {
                this.partCostRequestData.requestPartCostDetails[i].requestPartCostDetailCrCuReasons[
                  j
                ] = this.masterCrCuReasonsMap.get(idNo)
              }
            }
          }
          if (this.isRequestDraft()) {
            this.validateRows()
            this.validateRequestPartCost()
            this.masterRoutingForEmployeeCode(this.authClientService.getEmployeeCodeAuthed())
          } else {
            //add routing can rfi
            this.arrayEmployeeCodeCanRfi = []
            for (let i = 0; i < this.partCostRequestData.requestHistoryRoutings.length; i++) {
              if (
                this.partCostRequestData.requestHistoryRoutings[i].masterRoutingLevel <
                this.partCostRequestData.requestBase.masterRoutingLevel
              ) {
                this.arrayEmployeeCodeCanRfi.push(
                  this.partCostRequestData.requestHistoryRoutings[i].employeeCode
                )
              }
            }
          }
        } else {
          this.messageService.add({
            severity: 'error',
            summary: 'Error',
            detail: 'Error fetching part cost request'
          })
        }
      },
      error: (error) => {
        this.isLoading = false
        console.error('Error fetching part cost request:', error)
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'Error fetching part cost request'
        })
      }
    })
  }

  get isDiffCurrency(): boolean {
    let isDiffCurrency =
      this.partCostRequestData.requestPartCost.partMasterCurrencyId !=
      this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
    return isDiffCurrency
  }

  convertToUploadFiles(files: any = []) {
    var fileList: FileParameter[] = []
    for (let file of files) {
      fileList.push({
        data: file,
        fileName: file.name
      })
    }
    return fileList
  }

  getMasterRequestAction() {
    this.masterService.masterRequestAction().subscribe({
      next: (data: MasterRequestAction[]) => {
        this.masterRequestActions = data
        this.masterRequestActionsFilter = data.filter((x) => x.isActive == true)
        this.masterRequestActionsMap = new Map(
          data.map((d: MasterRequestAction) => {
            return [d.idNo, d]
          })
        )
        console.log('Master Request Action:', this.masterRequestActions)
      },
      error: (error) => {
        console.error('Error fetching master Request Action:', error)
      }
    })
  }
  getMasterRequestFilter() {
    this.masterService.masterRequestFilter().subscribe({
      next: (data: MasterRequestFilter[]) => {
        this.masterRequestFilters = data
        this.masterRequestFiltersFilter = data.filter((x) => x.isActive == true)
        this.masterRequestFiltersMap = new Map(
          data.map((d: MasterRequestFilter) => {
            return [d.idNo, d]
          })
        )
        console.log('Master Request Filter:', this.masterRequestFilters)
      },
      error: (error) => {
        console.error('Error fetching master Request Filter:', error)
      }
    })
  }

  isBuyMaterial() {
    return (
      this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId ==
      MasterTypeFormRequestPartCostEnum.BuyMaterial
    )
  }
  isBuyPart() {
    return (
      this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId ==
      MasterTypeFormRequestPartCostEnum.BuyPart
    )
  }

  getMasterRequestStatus() {
    this.masterService.masterRequestStatus().subscribe({
      next: (data: MasterRequestStatus[]) => {
        this.masterRequestStatuses = data
        this.masterRequestStatusesMap = new Map(
          data.map((d: MasterRequestStatus) => {
            return [d.idNo, d]
          })
        )
        console.log('Master Request Status:', this.masterRequestStatuses)
      },
      error: (error) => {
        console.error('Error fetching master Request Status:', error)
      }
    })
  }

  getMasterCrCuReason() {
    this.masterService.masterCrCuReason().subscribe({
      next: (data: MasterCrCuReason[]) => {
        this.masterCrCuReasons = data
        this.masterCrCuReasonsFilter = data.filter((x) => x.isActive == true)
        this.masterCrCuReasonsMap = new Map(
          data.map((d: MasterCrCuReason) => {
            return [d.idNo, d]
          })
        )
        console.log('Master MasterCrCuReason:', this.masterCurrencies)
      },
      error: (error) => {
        console.error('Error fetching master MasterCrCuReason:', error)
      }
    })
  }

  masterRoutingForEmployeeCode(employeeCode: string) {
    this.masterService.masterRoutingForEmployee(employeeCode).subscribe({
      next: (data: MasterRouting[]) => {
        this.partCostRequestData.masterRoutings = data
      },
      error: (error) => {}
    })
  }

  masterEmployee() {
    this.masterService.masterEmployee().subscribe({
      next: (data: MasterEmployee[]) => {
        this.masterEmployees = data
        this.masterEmployeesMap = new Map(
          data.map((d: MasterEmployee) => {
            return [d.employeeCode, d]
          })
        )
      },
      error: (error) => {}
    })
  }

  getMasterIncoterm() {
    this.masterService.masterIncoterm().subscribe({
      next: (data: MasterIncoterm[]) => {
        this.masterIncoterms = data
        this.masterIncotermsFilter = data.filter((x) => x.isActive == true)
        this.masterIncotermsMap = new Map(
          data.map((d: MasterIncoterm) => {
            return [d.idNo, d]
          })
        )
        console.log('MasterIncoterm:', this.masterIncoterms)
      },
      error: (error) => {
        console.error('Error fetching  masterIncoterms:', error)
      }
    })
  }

  getMasterOriginal() {
    this.masterService.masterOriginal().subscribe({
      next: (data: MasterOriginal[]) => {
        this.masterOriginals = data
        this.masterOriginalsFilter = data.filter((x) => x.isActive == true)
        this.masterOriginalsMap = new Map(
          data.map((d: MasterOriginal) => {
            return [d.idNo, d]
          })
        )
        console.log('Master Original:', this.masterOriginals)
      },
      error: (error) => {
        console.error('Error fetching master Original:', error)
      }
    })
  }

  initPartCostRequestData() {
    this.partCostRequestData = this.helperService.createNew(PartCostRequestData)
    this.partCostRequestData.masterRoutings = []
    this.partCostRequestData.requestApprovalSummaries = []
    this.partCostRequestData.requestBase = this.helperService.createNew(RequestBase)
    this.partCostRequestData.requestBase.masterRequestStatusId = this.helperService.intMin
    this.partCostRequestData.requestBase.masterRoutingLevel = 1
    this.partCostRequestData.requestBase.masterRequestTypeId = this.helperService.intMin
    this.partCostRequestData.requestBase.requestNo = this.helperService.intMin
    this.partCostRequestData.requestBaseSupportDocuments = []
    this.partCostRequestData.requestBaseQuotationDocuments = []
    this.partCostRequestData.requestHistoryRoutings = []
    this.partCostRequestData.masterChangePriceReasonLogs = []
    this.partCostRequestData.requestPartCost = this.helperService.createNew(RequestPartCost)
    this.partCostRequestData.requestPartCost.buyFrom = 'PAPVN'
    this.partCostRequestData.requestPartCost.unit = 'KG'
    this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId =
      MasterTypeRequestPartCostEnum.Normal
    this.partCostRequestData.requestPartCost.requestBaseId = this.partCostRequestData.requestBase.id
    this.partCostRequestData.requestPartCost.requestBaseIdNo =
      this.partCostRequestData.requestBase.idNo
    this.partCostRequestData.requestPartCostDetails = []
    this.validateRequestPartCost()
  }

  getMasterSupplier() {
    this.masterService.masterSupplier().subscribe({
      next: (data: MasterSupplier[]) => {
        this.masterSuppliers = data
        this.masterSuppliersFilter = data.filter((x) => x.isActive == true)
        this.masterSuppliersMap = new Map(
          data.map((d: MasterSupplier) => {
            return [d.supplierCode, d]
          })
        )
        console.log('Master Suppliers:', this.masterSuppliers)
      },
      error: (error) => {
        console.error('Error fetching master suppliers:', error)
      }
    })
  }

  getMasterCurrency() {
    this.masterService.masterCurrency().subscribe({
      next: (data: MasterCurrency[]) => {
        this.masterCurrencies = data
        this.masterCurrenciesFilter = data.filter((x) => x.isActive == true)
        this.masterCurrenciesMap = new Map(
          data.map((d: MasterCurrency) => {
            return [d.idNo, d]
          })
        )
        console.log('Master masterCurrencies:', this.masterCurrencies)
      },
      error: (error) => {
        console.error('Error fetching master masterCurrencies:', error)
      }
    })
  }

  getMasterTypeRequestPartCost() {
    this.masterService.masterTypeRequestPartCost().subscribe({
      next: (data: MasterTypeRequestPartCost[]) => {
        this.masterTypeRequestPartCosts = data
        this.masterTypeRequestPartCostsMap = new Map(
          data.map((d: MasterTypeRequestPartCost) => {
            return [d.idNo, d]
          })
        )
        console.log('Master MasterTypeRequestPartCost:', this.masterTypeRequestPartCosts)
      },
      error: (error) => {
        console.error('Error fetching master MasterTypeRequestPartCost:', error)
      }
    })
  }

  getMasterChangePriceReason() {
    this.masterService.masterChangePriceReason().subscribe({
      next: (data: MasterChangePriceReason[]) => {
        this.masterChangePriceReasons = data
        this.masterChangePriceReasonsFilter = data.filter((x) => x.isActive == true)
        this.masterChangePriceReasonsMap = new Map(
          data.map((d: MasterChangePriceReason) => {
            return [d.idNo, d]
          })
        )
        console.log('Master Change Price Reasons:', this.masterChangePriceReasons)
      }
    })
  }
  getMasterTypeFormRequestPartCost() {
    this.masterService.masterTypeFormRequestPartCost().subscribe({
      next: (data: MasterTypeFormRequestPartCost[]) => {
        this.masterTypeFormRequestPartCosts = data
        this.masterTypeFormRequestPartCostsMap = new Map(
          data.map((d: MasterTypeFormRequestPartCost) => {
            return [d.idNo, d]
          })
        )
        console.log('MasterTypeFormRequestPartCost:', this.masterChangePriceReasons)
      }
    })
  }

  customSupplierAliasFn(masterSupplier: MasterSupplier) {
    return `${masterSupplier.supplierCode} - ${masterSupplier.supplierAlias} - ${masterSupplier.supplierName}`
  }

  customSearchFromVendorCodeFn(term: string, item: MasterSupplier): boolean {
    term = term.toLowerCase()
    return (
      item.supplierCode.toLowerCase().includes(term) ||
      item.supplierAlias.toLowerCase().includes(term) ||
      item.supplierName.toLowerCase().includes(term)
    )
  }
  customSearchToVendorCodeFn(term: string, item: MasterSupplier): boolean {
    term = term.toLowerCase()
    return (
      item.supplierCode.toLowerCase().includes(term) ||
      item.supplierAlias.toLowerCase().includes(term) ||
      item.supplierName.toLowerCase().includes(term)
    )
  }

  onUpload(event: any) {
    if (event.files && event.files.length > 0) {
      for (let file of event.files) {
        this.uploadedFiles.push(file) // Add files to the temporary array
      }
      this.messageService.add({
        severity: 'success',
        summary: 'File Uploaded',
        detail: 'Files added successfully.'
      })
    }
  }

  deleteUploadedFile(index: number) {
    if (index >= 0 && index < this.uploadedFiles.length) {
      const removedFile = this.uploadedFiles.splice(index, 1) // Remove the file from the temporary array
      this.messageService.add({
        severity: 'info',
        summary: 'File Deleted',
        detail: `${removedFile[0].name} removed.`
      })
    }
  }

  onUploadSelect(event: any) {
    for (let file of event.files) {
      if (this.uploadedFiles.find((f) => f.name === file.name)) {
        this.messageService.add({
          severity: 'warn',
          summary: 'File Already Exists',
          detail: `${file.name} already exists.`
        })
        return
      } else {
        this.uploadedFiles.push(file) // Add files to the temporary array
      }
    }
  }
  onUploadRemove(event: any) {
    console.log('onRemove', event)
    const index = this.uploadedFiles.findIndex((file) => file.name === event.file.name)
    if (index !== -1) {
      this.uploadedFiles.splice(index, 1)
    }
  }
  onRemoveUploadedFile(event: any) {
    console.log('onRemoveUploadedFile', event)
  }
  onClear(event: any) {
    console.log('onClear', event)
  }
  onUploadQuotationSelect(event: any) {
    for (let file of event.files) {
      if (this.uploadedQuotationFiles.find((f) => f.name === file.name)) {
        this.messageService.add({
          severity: 'warn',
          summary: 'File Already Exists',
          detail: `${file.name} already exists.`
        })
        return
      } else {
        this.uploadedQuotationFiles.push(file) // Add files to the temporary array
      }
    }
  }
  onUploadQuotationRemove(event: any) {
    console.log('onRemove', event)
    const index = this.uploadedFiles.findIndex((file) => file.name === event.file.name)
    if (index !== -1) {
      this.uploadedQuotationFiles.splice(index, 1)
    }
  }
  // onRemoveUploadedQuotationFile(event: any) {
  //   console.log('onRemoveUploadedFile', event)
  // }
  // onQuotationClear(event: any) {
  //   console.log('onClear', event)
  // }

  // Enhanced version to avoid unnecessary calculations
  validateRow(partCost: RequestPartCostDetail): boolean {
    let isValid = true

    // Reset errors for this specific row
    this.errors[partCost.id as string] = {}

    // Basic information validation - always required
    if (partCost.isParentPart === undefined) {
      this.errors[partCost.id as string]['isParentPart'] = true
      isValid = false
    }

    if (!partCost.partNameSpec || partCost.partNameSpec.trim() === '') {
      this.errors[partCost.id as string]['partNameSpec'] = true
      isValid = false
    }

    // Present values validation - allow 0 but not negative or null/undefined
    if (
      (partCost.presentTotalPrice === null ||
       partCost.presentTotalPrice === undefined ||
       partCost.presentTotalPrice < 0) &&
      partCost.isNewPart == false
    ) {
      this.errors[partCost.id as string]['presentTotalPrice'] = true
      isValid = false
    }

    // New values validation - allow 0 but not negative or null/undefined
    if (partCost.newTotalPrice === null ||
        partCost.newTotalPrice === undefined ||
        partCost.newTotalPrice < 0) {
      this.errors[partCost.id as string]['newTotalPrice'] = true
      isValid = false
    }

    if (partCost.basePartCode) {
      partCost.basePartCode = partCost.basePartCode.trim()
    }

    if (partCost.newPartCode) {
      partCost.newPartCode = partCost.newPartCode.trim()
    }

    // Conditional validation based on isParent
    if (partCost.isParentPart == true) {
      if (!partCost.newPartCode) {
        this.errors[partCost.id as string]['newPartCode'] = true
        isValid = false
      }
      if (!partCost.leadTimeDay) {
        this.errors[partCost.id as string]['leadTimeDay'] = true
        isValid = false
      }
      if (this.isDiffCurrency == true) {
        if (
          !partCost.presentExchangeRate ||
          partCost.presentExchangeRate <= 0 ||
          partCost.presentExchangeRate == null
        ) {
          this.errors[partCost.id as string]['presentExchangeRate'] = true
          isValid = false
        }
        if (
          !partCost.newExchangeRate ||
          partCost.newExchangeRate <= 0 ||
          partCost.newExchangeRate == null
        ) {
          this.errors[partCost.id as string]['newExchangeRate'] = true
          isValid = false
        }
      }

      // Always required fields regardless of isParent
      if (!partCost.effectiveDateFrom) {
        this.errors[partCost.id as string]['effectiveDateFrom'] = true
        isValid = false
      } else {
        if (
          this.dateHelperService.isBefore(partCost.effectiveDateFrom, this.firstSubmitDate) &&
          this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId !=
            MasterTypeRequestPartCostEnum.BackDate
        ) {
          this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId =
            MasterTypeRequestPartCostEnum.BackDate
        }
      }

      if (!partCost.effectiveDateTo) {
        this.errors[partCost.id as string]['effectiveDateTo'] = true
        isValid = false
      }
      // if (!partCost.reason || partCost.reason.trim() === '') {
      //   this.errors[partCost.id as string]['reason'] = true
      //   isValid = false
      // }

      if (partCost.requestPartCostDetailCrCuReasons.length === 0 && partCost.isNewPart == false) {
        this.errors[partCost.id as string]['requestPartCostDetailCrCuReasons'] = true
        isValid = false
      }

      if (!partCost.masterOriginalId) {
        this.errors[partCost.id as string]['masterOriginalId'] = true
        isValid = false
      }

      if (
        (partCost.presentMoq === undefined || partCost.presentMoq === null) &&
        partCost.isNewPart == false
      ) {
        this.errors[partCost.id as string]['presentMoq'] = true
        isValid = false
      }
      if (
        (partCost.presentTotalPrice === null ||
         partCost.presentTotalPrice === undefined ||
         partCost.presentTotalPrice < 0) &&
        partCost.isNewPart == false
      ) {
        this.errors[partCost.id as string]['presentTotalPrice'] = true
        isValid = false
      }

      if (
        (partCost.presentIncotermId === undefined || partCost.presentIncotermId === null) &&
        partCost.isNewPart == false
      ) {
        this.errors[partCost.id as string]['presentIncotermId'] = true
        isValid = false
      }

      if (
        (partCost.presentImportDuty === undefined || partCost.presentImportDuty === null) &&
        partCost.isNewPart == false
      ) {
        this.errors[partCost.id as string]['presentImportDuty'] = true
        isValid = false
      }
      if (partCost.newMoq === undefined || partCost.newMoq === null) {
        this.errors[partCost.id as string]['newMoq'] = true
        isValid = false
      }

      if (partCost.newIncotermId === undefined || partCost.newIncotermId === null) {
        this.errors[partCost.id as string]['newIncotermId'] = true
        isValid = false
      }

      if (partCost.newImportDuty === undefined || partCost.newImportDuty === null) {
        this.errors[partCost.id as string]['newImportDuty'] = true
        isValid = false
      }
      if (!partCost.newTotalPrice || partCost.newTotalPrice < 0) {
        this.errors[partCost.id as string]['newTotalPrice'] = true
        isValid = false
      }
    }
    return isValid
  }

  // Check if a field is invalid
  isInvalid(partCost: RequestPartCostDetail, field: string): boolean {
    // If the row doesn't have any errors yet, initialize an empty object
    if (!this.errors[partCost.id as string]) {
      this.errors[partCost.id as string] = {}
    }
    // For all other fields
    let result =
      (this.submitted || this.errors[partCost.id as string][field]) &&
      this.errors[partCost.id as string][field]
    return result
  }

  addNewRow() {
    let newRow = this.helperService.createNew(RequestPartCostDetail)
    newRow.id = this.helperService.generateUUID()
    newRow.isParentPart = true
    newRow.isUploadSap = true
    newRow.isNewPart = false
    newRow.requestBaseId = this.partCostRequestData.requestBase.id
    newRow.requestPartCostId = this.partCostRequestData.requestPartCost.id
    newRow.orderNo = this.partCostRequestData.requestPartCostDetails.length + 1
    newRow.requestPartCostDetailCrCuReasons = []
    if (this.partCostRequestData.requestPartCost.isSameLeadTime == true) {
      newRow.leadTimeDay = this.partCostRequestData.requestPartCost.leadTimeDay
    }
    this.partCostRequestData.requestPartCostDetails.push(newRow)
    this.validateRow(newRow)
    setTimeout(() => this.editMode(), 500)
  }

  titleTable(): string {
    return this.isFullScreen ? 'Zoom Out' : 'Zoom In'
  }

  copyNewRow() {
    let newRowCopy =
      this.partCostRequestData.requestPartCostDetails[
        this.partCostRequestData.requestPartCostDetails.length - 1
      ]
    let newRow = this.helperService.copyItem(RequestPartCostDetail, newRowCopy)
    newRow.id = this.helperService.generateUUID()
    newRow.orderNo = this.partCostRequestData.requestPartCostDetails.length
    // console.log('newRow', newRow)

    this.partCostRequestData.requestPartCostDetails.push(newRow)
    this.validateRow(newRow)
    setTimeout(() => this.editMode(), 500)
  }

  validateRows(): boolean {
    let isValidRows = true
    this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId =
      MasterTypeRequestPartCostEnum.Normal
    this.partCostRequestData.requestPartCostDetails.forEach((partCost) => {
      let isValidRow = this.validateRow(partCost)
      if (isValidRow == false) {
        isValidRows = isValidRow
      }
    })
    return isValidRows
  }

  deleteRow(index: number) {
    this.partCostRequestData.requestPartCostDetails.splice(index, 1)
  }

  isDisableInputPresentTotalPrice(partCost: RequestPartCostDetail) {
    let result = (
      (((partCost.presentMaterialCost && partCost.presentMaterialCost > 0) ||
        (partCost.presentProcessingCost && partCost.presentProcessingCost > 0) ||
        (partCost.presentOtherCost && partCost.presentOtherCost > 0)) &&
        this.isBuyPart()) ||
      this.isBuyMaterial()
    )
          return result

  }
isDisableInputNewTotalPrice(partCost: RequestPartCostDetail) {
    return (
      (((partCost.newMaterialCost && partCost.newMaterialCost > 0) ||
        (partCost.newProcessingCost && partCost.newProcessingCost > 0) ||
        (partCost.newOtherCost && partCost.newOtherCost > 0)) &&
        this.isBuyPart()) ||
      this.isBuyMaterial()
    )
  }

  isInvalidRequestPartCost(field: string): boolean {
    return this.errorsGeneral[field]
  }

  isValidGeneral(): boolean {
    let isValid = true
    // Reset errors for general fields
    this.errorsGeneral = {}

    // Validate masterTypeFormRequestPartCostId
    if (!this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId) {
      this.errorsGeneral['masterTypeFormRequestPartCostId'] = true
      isValid = false
    }

    // Validate masterTypeRequestPartCostId
    if (!this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId) {
      this.errorsGeneral['masterTypeRequestPartCostId'] = true
      isValid = false
    }

    // Validate partMasterCurrencyId
    if (!this.partCostRequestData.requestPartCost.partMasterCurrencyId) {
      this.errorsGeneral['partMasterCurrencyId'] = true
      isValid = false
    }

    // Validate paymentMasterCurrencyId
    if (!this.partCostRequestData.requestPartCost.paymentMasterCurrencyId) {
      this.errorsGeneral['paymentMasterCurrencyId'] = true
      isValid = false
    }
    if (!this.partCostRequestData.requestPartCost.fromVendorCode) {
      this.errorsGeneral['fromVendorCode'] = true
      isValid = false
    }
    if (!this.partCostRequestData.requestPartCost.toVendorCode) {
      this.errorsGeneral['toVendorCode'] = true
      isValid = false
    }
    if (this.partCostRequestData.masterChangePriceReasonLogs.length == 0) {
      this.errorsGeneral['masterChangePriceReasonLogs'] = true
      isValid = false
    }

    if (
      !this.partCostRequestData.requestBase.shortDescription ||
      this.partCostRequestData.requestBase.shortDescription == null ||
      this.partCostRequestData.requestBase.shortDescription == ''
    ) {
      this.errorsGeneral['shortDescription'] = true
      isValid = false
    }

    if (
      !this.partCostRequestData.requestPartCost.remark ||
      this.partCostRequestData.requestPartCost.remark == null ||
      this.partCostRequestData.requestPartCost.remark == ''
    ) {
      this.errorsGeneral['remark'] = true
      isValid = false
    }
    return isValid
  }

  filterMasterActived() {
    if (this.isDisableInput()) {
      this.masterSuppliersFilter = [...this.masterSuppliers]
      this.masterCrCuReasonsFilter = [...this.masterCrCuReasons]
      // this.masterTypeRequestPartCostsFilter = [...this.masterTypeRequestPartCosts]
      // this.masterTypeFormRequestPartCostsFilter = [...this.masterTypeFormRequestPartCosts]
      this.masterIncotermsFilter = [...this.masterIncoterms]
      this.masterCurrenciesFilter = [...this.masterCurrencies]
      this.masterOriginalsFilter = [...this.masterOriginals]
      this.masterRequestFiltersFilter = [...this.masterRequestFilters]
    } else {
      this.masterSuppliersFilter = this.masterSuppliers.filter((x) => x.isActive == true)
      this.masterCrCuReasonsFilter = this.masterCrCuReasons.filter((x) => x.isActive == true)
      this.masterIncotermsFilter = this.masterIncoterms.filter((x) => x.isActive == true)
      this.masterCurrenciesFilter = this.masterCurrencies.filter((x) => x.isActive == true)
      this.masterOriginalsFilter = this.masterOriginals.filter((x) => x.isActive == true)
    }
  }

  validateRequestPartCost(): boolean {
    let isValid = true
    // Reset errors for general fields
    this.errorsGeneral = {}

    isValid = this.isValidGeneral()

    if (
      !this.partCostRequestData.requestPartCost.remark ||
      this.partCostRequestData.requestPartCost.remark == null ||
      this.partCostRequestData.requestPartCost.remark == ''
    ) {
      this.errorsGeneral['remark'] = true
      isValid = false
    }
    if (
      this.isBuyMaterial() &&
      (!this.partCostRequestData.requestPartCost.unit ||
        this.partCostRequestData.requestPartCost.unit == null ||
        this.partCostRequestData.requestPartCost.unit == '')
    ) {
      this.errorsGeneral['unit'] = true
      isValid = false
    }

    return isValid
  }

  changeVendor() {
    console.log('changeVendor', this.partCostRequestData.requestPartCost.toVendorCode)
    if (this.partCostRequestData.requestPartCost.toVendorCode) {
      let toVendor = this.masterSuppliersMap.has(
        this.partCostRequestData.requestPartCost.toVendorCode
      )
        ? this.masterSuppliersMap.get(this.partCostRequestData.requestPartCost.toVendorCode)
        : null

      console.log('changeVendor', toVendor)

      if (toVendor) {
        let isValidTypeForm = this.masterTypeFormRequestPartCostsMap.has(
          toVendor.masterTypeFormRequestPartCostId
        )
        let isValidMasterCurrencyPartId = this.masterCurrenciesMap.has(
          toVendor.masterCurrencyPartId
        )
        let isValidMasterCurrencyPaymentId = this.masterCurrenciesMap.has(
          toVendor.masterCurrencyPaymentId
        )

        if (isValidTypeForm) {
          this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId =
            toVendor.masterTypeFormRequestPartCostId
        } else {
          this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId = null
        }
        if (isValidMasterCurrencyPartId) {
          this.partCostRequestData.requestPartCost.paymentMasterCurrencyId =
            toVendor.masterCurrencyPartId
          this.currencyPart = this.masterCurrenciesMap.get(toVendor.masterCurrencyPartId)
        } else {
          this.partCostRequestData.requestPartCost.paymentMasterCurrencyId = null
          this.currencyPart = null
        }
        if (isValidMasterCurrencyPaymentId) {
          this.partCostRequestData.requestPartCost.partMasterCurrencyId =
            toVendor.masterCurrencyPaymentId
          this.currencyPayment = this.masterCurrenciesMap.get(toVendor.masterCurrencyPaymentId)
        } else {
          this.partCostRequestData.requestPartCost.partMasterCurrencyId = null
          this.currencyPayment = null
        }

        this.changePartCurrency()
        this.changePaymentCurrency()
      }
    }
  }

  isPendingApprover() {
    return (
      this.authClientService.getEmployeeCodeAuthed() ==
        this.partCostRequestData.requestBase.pendingBy &&
      this.partCostRequestData.requestBase.masterRequestStatusId != StatusRequestEnum.Draft
    )
  }

  isDisableInput(): boolean {
    if (this.isView) {
      if (this.isRequestDraft()) {
        return false
      } else {
        return true
      }
    } else return false
  }

  isRequestDraft() {
    return (
      this.authClientService.getEmployeeCodeAuthed() ==
        this.partCostRequestData.requestBase.pendingBy &&
      this.partCostRequestData.requestBase.masterRequestStatusId == StatusRequestEnum.Draft
    )
  }
  isRequestorCanGetBack() {
    return (
      this.authClientService.getEmployeeCodeAuthed() ==
        this.partCostRequestData.requestBase.requestBy &&
      this.partCostRequestData.requestBase.masterRequestStatusId ==
        StatusRequestEnum.PendingFirstApprover
    )
  }

  isTopUpRole() {
    return (
      this.authClientService.isTopUpRole() &&
      this.partCostRequestData.requestBase.masterRequestStatusId != StatusRequestEnum.Draft
    )
  }

  isRequestPendingApprover() {
    return (
      (this.partCostRequestData.requestBase.masterRequestStatusId ==
        StatusRequestEnum.PendingApprover ||
        this.partCostRequestData.requestBase.masterRequestStatusId ==
          StatusRequestEnum.PendingFirstApprover) &&
      this.partCostRequestData.requestBase.pendingBy ==
        this.authClientService.getEmployeeCodeAuthed()
    )
  }

  approveRequestPartCost() {
    this.submitted = true
    let isValid = true
    this.isRequiredRemark = false
    this.actionData.actionId = ActionRequestEnum.Approve
    // this.actionData.requestPartCostDetails = this.partCostRequestData.requestPartCostDetails
    if (isValid) {
      this.isLoading = true
      this.partCostService.approveRequestPartCost(this.actionData).subscribe({
        next: (data: ResponseData) => {
          this.isLoading = false

          if (data.isError == false) {
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'Request approved successfully'
            })
            // let linkRedirect =
            //   '/request_part_cost/view/' + this.partCostRequestData.requestBase.requestNo
            // this.router.navigate([linkRedirect])
            location.reload()
          } else {
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error'
            })
          }
        }
      })
    } else {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }
  rejectRequestPartCost() {
    this.submitted = true
    let isValid = true
    this.actionData.actionId = ActionRequestEnum.Reject
    this.actionData.requestPartCostDetails = this.partCostRequestData.requestPartCostDetails
    let isCommentAtEachRow = false
    for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
      if (
        this.partCostRequestData.requestPartCostDetails[i].approverComment != undefined &&
        this.partCostRequestData.requestPartCostDetails[i].approverComment?.trim() != ''
      ) {
        isCommentAtEachRow = true
        break
      }
    }
    if (
      isCommentAtEachRow == false &&
      (this.actionData.comment == '' || this.actionData.comment == null)
    ) {
      this.isRequiredRemark = true
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please input comment when Reject'
      })
      return
    } else {
      this.isRequiredRemark = false
    }
    if (isValid) {
      this.isLoading = true
      this.partCostService.rejectRequestPartCost(this.actionData).subscribe({
        next: (data: ResponseData) => {
          this.isLoading = false

          if (data.isError == false) {
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'Request Reject successfully'
            })
            // let linkRedirect =
            //   '/request_part_cost/view/' + this.partCostRequestData.requestBase.requestNo
            // this.router.navigate([linkRedirect])
            location.reload()
          } else {
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error'
            })
          }
        }
      })
    } else {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }

  rfiRequestPartCost(isRfiToPic: boolean = false) {
    this.submitted = true
    let isValid = true
    this.actionData.actionId = ActionRequestEnum.RFI
    this.actionData.requestPartCostDetails = this.partCostRequestData.requestPartCostDetails
    let isCommentAtEachRow = false
    for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
      if (
        this.partCostRequestData.requestPartCostDetails[i].approverComment != undefined &&
        this.partCostRequestData.requestPartCostDetails[i].approverComment?.trim() != ''
      ) {
        isCommentAtEachRow = true
        break
      }
    }
    if (
      isCommentAtEachRow == false &&
      (this.actionData.comment == '' || this.actionData.comment == null)
    ) {
      this.isRequiredRemark = true
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please input comment when RFI'
      })
      return
    } else {
      this.isRequiredRemark = false
    }
    if (isRfiToPic == true) {
      this.actionData.employeeCodeRfi = this.partCostRequestData.requestBase.requestBy
      this.actionData.masterRoutingLevel = 0
    } else if (!this.actionData.employeeCodeRfi) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please choose employee when RFI'
      })
      return
    }
    if (isValid) {
      this.isLoading = true
      this.partCostService.rfiRequestPartCost(this.actionData).subscribe({
        next: (data: ResponseData) => {
          this.isLoading = false

          if (data.isError == false) {
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'Request Reject successfully'
            })
            this.isLoading = true
            // let linkRedirect =
            //   '/request_part_cost/view/' + this.partCostRequestData.requestBase.requestNo
            // this.router.navigate([linkRedirect])
            location.reload()
          } else {
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error'
            })
          }
        }
      })
    } else {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }
  getBackRequestPartCost() {
    this.submitted = true
    let isValid = true
    this.actionData.actionId = ActionRequestEnum.GetBack
    this.actionData.requestPartCostDetails = this.partCostRequestData.requestPartCostDetails
    if (this.actionData.comment == '' || this.actionData.comment == null) {
      this.isRequiredRemark = true
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please input comment when GetBack'
      })
      return
    }
    if (isValid) {
      this.isLoading = true
      this.partCostService.getBackRequestPartCost(this.actionData).subscribe({
        next: (data: ResponseData) => {
          this.isLoading = false

          if (data.isError == false) {
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'Request GetBack successfully'
            })
            // let linkRedirect =
            //   '/request_part_cost/view/' + this.partCostRequestData.requestBase.requestNo
            // this.router.navigate([linkRedirect])
            location.reload()
          } else {
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error'
            })
          }
        }
      })
    } else {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }
  topUpRequestPartCost() {
    this.submitted = true
    let isValid = true
    this.actionData.actionId = ActionRequestEnum.TopUp
    this.actionData.requestPartCostDetails = this.partCostRequestData.requestPartCostDetails
    let isCommentAtEachRow = false
    for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
      if (
        this.partCostRequestData.requestPartCostDetails[i].approverComment != undefined &&
        this.partCostRequestData.requestPartCostDetails[i].approverComment?.trim() != ''
      ) {
        isCommentAtEachRow = true


        break
      }
    }
    if (
      isCommentAtEachRow == false &&
      (this.actionData.comment == '' || this.actionData.comment == null)
    ) {
      this.isRequiredRemark = true
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please input comment when TopUp'
      })
      return
    }
    if (isValid) {
      this.isLoading = true
      this.partCostService.topUpRequestPartCost(this.actionData).subscribe({
        next: (data: ResponseData) => {
          this.isLoading = false

          if (data.isError == false) {
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'Request GetBackSpecial successfully'
            })
            let linkRedirect =
              '/request_part_cost/view/' + this.partCostRequestData.requestBase.requestNo
            this.router.navigate([linkRedirect])
            location.reload()
          } else {
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error'
            })
          }
        }
      })
    } else {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }

  createRequestPartCost(isResubmit: boolean = false, isSaveDraft: boolean = false) {
    this.submitted = true
    let isValid = true
    // Validate all rows
    isValid = this.validateRows()

    if (isValid == true) {
      isValid = this.validateRequestPartCost()
    }

    if (isValid) {
      this.isLoading = true
      this.partCostService
        .createRequestPartCost(
          isResubmit,
          this.actionData.comment ?? '',
          isSaveDraft,
          this.partCostRequestData
        )
        .subscribe({
          next: (data: ResponseData) => {
            if (data.isError == false) {
              let isNeedUploadFile =
                this.uploadedFiles.length > 0 || this.uploadedQuotationFiles.length > 0
              let linkRedirect = '/request_part_cost/view/' + data.message
              if (isNeedUploadFile) {
                this.partCostService
                  .updateRequestPartCostWithFile(
                    this.convertToUploadFiles(this.uploadedFiles),
                    this.convertToUploadFiles(this.uploadedQuotationFiles),
                    data.message,
                    []
                  )
                  .subscribe({
                    next: (responseUpload: ResponseData) => {
                      this.messageService.add({
                        severity: 'success',
                        summary: 'Success',
                        detail: 'Successfully'
                      })
                      if (this.isView) {
                        location.reload()
                      } else {
                        this.router.navigate([linkRedirect])
                      }
                    },
                    error: (errorUpload) => {
                      this.isLoading = false
                      this.messageService.add({
                        severity: 'error',
                        summary: 'Error',
                        detail: 'Upload file error! Please re-upload file.'
                      })
                    }
                  })
              } else {
                // this.isLoading = false
                this.messageService.add({
                  severity: 'success',
                  summary: 'Success',
                  detail: 'Successfully'
                })
                if (this.isView) {
                  location.reload()
                } else {
                  this.router.navigate([linkRedirect])
                }
              }
            } else {
              this.isLoading = false
              this.messageService.add({
                severity: 'error',
                summary: 'Error',
                detail: 'Please fill all required fields'
              })
            }
          },
          error: (error) => {
            this.isLoading = false
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: error.message
            })
          }
        })
    } else {
      this.isLoading = false

      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please fill all required fields'
      })
    }
  }
  // Export to Excel with columns matching the table structure
  exportToExcel(): void {
    const workbook = XLSX.utils.book_new()

    // Style definitions
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: 'E6F3FF' }, patternType: 'solid' }
    }

    const firstColumnStyle = {
      fill: { fgColor: { rgb: 'E6F3FF' }, patternType: 'solid' }
    }

    // Helper function to apply styles to a worksheet
    const applyStyles = (worksheet: XLSX.WorkSheet) => {
      const range = XLSX.utils.decode_range(worksheet['!ref'])

      // Style headers (first row)
      for (let C = range.s.c; C <= range.e.c; C++) {
        const headerCell = XLSX.utils.encode_cell({ r: 0, c: C })
        if (!worksheet[headerCell].s) worksheet[headerCell].s = {}
        Object.assign(worksheet[headerCell].s, headerStyle)
      }

      // Style first column
      for (let R = range.s.r; R <= range.e.r; R++) {
        const firstColCell = XLSX.utils.encode_cell({ r: R, c: 0 })
        if (!worksheet[firstColCell].s) worksheet[firstColCell].s = {}
        Object.assign(worksheet[firstColCell].s, firstColumnStyle)
      }
    }

    // Add and style main data sheet
    const worksheet = XLSX.utils.aoa_to_sheet([
      Object.values(this.createExportHeaders()),
      ...this.partCostRequestData.requestPartCostDetails.map((partCost) =>
        Object.values(this.mapPartCostToExportRow(partCost))
      )
    ])
    applyStyles(worksheet)
    XLSX.utils.book_append_sheet(workbook, worksheet, 'PartCostDetails')

    // Add and style Currency sheet
    const currencyHeaders = ['Name']
    const currencyData = [
      currencyHeaders,
      ...this.masterCurrenciesFilter.map((item) => [item.name])
    ]
    const currencySheet = XLSX.utils.aoa_to_sheet(currencyData)
    applyStyles(currencySheet)
    XLSX.utils.book_append_sheet(workbook, currencySheet, 'Currencies')

    // Add and style Incoterms sheet
    const incotermsHeaders = ['Name']
    const incotermsData = [
      incotermsHeaders,
      ...this.masterIncotermsFilter.map((item) => [item.name])
    ]
    const incotermsSheet = XLSX.utils.aoa_to_sheet(incotermsData)
    applyStyles(incotermsSheet)
    XLSX.utils.book_append_sheet(workbook, incotermsSheet, 'Incoterms')

    // Add and style CU-CR Reasons sheet
    const cucrHeaders = ['Definition']
    const cucrData = [
      cucrHeaders,
      ...this.masterCrCuReasonsFilter.map((item) => [item.masterDefinition])
    ]
    const cucrSheet = XLSX.utils.aoa_to_sheet(cucrData)
    applyStyles(cucrSheet)
    XLSX.utils.book_append_sheet(workbook, cucrSheet, 'CU-CR_Reasons')

    // Add and style Originals sheet
    const originalHeaders = ['Name']
    const originalData = [originalHeaders, ...this.masterOriginalsFilter.map((item) => [item.name])]
    const originalSheet = XLSX.utils.aoa_to_sheet(originalData)
    applyStyles(originalSheet)
    XLSX.utils.book_append_sheet(workbook, originalSheet, 'Originals')

    // Export workbook
    const excelBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
      cellStyles: true
    })

    const data = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })

    const url = URL.createObjectURL(data)
    const link = document.createElement('a')
    link.href = url
    link.download = 'PartCostDetails.xlsx'
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  // Helper function to create headers for Excel export
  private createExportHeaders(): { [key: string]: string } {
    let columnIndex = 0
    const headers: { [key: string]: string } = {}

    // Basic Information - match exact column order from HTML table
    headers[this.toColumnLetter(columnIndex++)] = 'Is Parent Part'
    headers[this.toColumnLetter(columnIndex++)] = 'Is Upload SAP'
    headers[this.toColumnLetter(columnIndex++)] = 'Is New Part'
    headers[this.toColumnLetter(columnIndex++)] = 'Base Part Code'
    headers[this.toColumnLetter(columnIndex++)] = 'New Part Code'
    headers[this.toColumnLetter(columnIndex++)] = 'Part Name/Spec'

    // Present values section - must match table columns
    const isBuyPart = this.isBuyPart()
    const currencyPart = this.getCurrencyName(      this.partCostRequestData.requestPartCost.partMasterCurrencyId
    )
    const currencyPayment = this.getCurrencyName(
      this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
    )
    const unit = this.partCostRequestData.requestPartCost.unit || 'KG'

    if (isBuyPart) {
      headers[this.toColumnLetter(columnIndex++)] = 'Present Material Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'Present Processing Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'Present Other Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'Present Total Price'
    } else {
      headers[this.toColumnLetter(columnIndex++)] = `Present ${currencyPart}/KG`
      headers[this.toColumnLetter(columnIndex++)] = 'Present Weight/pc'
      headers[this.toColumnLetter(columnIndex++)] = 'Present Other Cost'
      headers[this.toColumnLetter(columnIndex++)] = `Present Total Price (${currencyPart}/pcs)`
    }

    if (this.isDiffCurrency) {
      headers[this.toColumnLetter(columnIndex++)] = 'Present Exchange Rate'
      headers[this.toColumnLetter(columnIndex++)] = `Present Total Price (${currencyPayment}/pcs)`
    }

    headers[this.toColumnLetter(columnIndex++)] = 'Present MOQ'
    headers[this.toColumnLetter(columnIndex++)] = 'Present Incoterm'
    headers[this.toColumnLetter(columnIndex++)] = 'Present Import Duty'
    headers[this.toColumnLetter(columnIndex++)] = 'Present Currency'    // New values section
    if (isBuyPart) {
      headers[this.toColumnLetter(columnIndex++)] = 'New Material Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'New Processing Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'New Other Cost'
      headers[this.toColumnLetter(columnIndex++)] = 'New Total Price'
    } else {
      headers[this.toColumnLetter(columnIndex++)] = `New ${currencyPart}/KG`
      headers[this.toColumnLetter(columnIndex++)] = 'New Weight/pc'
      headers[this.toColumnLetter(columnIndex++)] = 'New Other Cost'
      headers[this.toColumnLetter(columnIndex++)] = `New Total Price (${currencyPart}/pcs)`
    }

    // New Other Cost - common for both Buy Part and Buy Material


    if (this.isDiffCurrency) {
      headers[this.toColumnLetter(columnIndex++)] = 'New Exchange Rate'
      headers[this.toColumnLetter(columnIndex++)] = `New Total Price (${currencyPayment}/pcs)`
    }

    headers[this.toColumnLetter(columnIndex++)] = 'New MOQ'
    headers[this.toColumnLetter(columnIndex++)] = 'New Incoterm'
    headers[this.toColumnLetter(columnIndex++)] = 'New Import Duty'
    headers[this.toColumnLetter(columnIndex++)] = 'New Currency'

    // Difference columns
    if (this.isDiffCurrency) {
      headers[this.toColumnLetter(columnIndex++)] = `Diff ${currencyPart}`
      headers[this.toColumnLetter(columnIndex++)] = `Diff % ${currencyPart}`
    }

    headers[this.toColumnLetter(columnIndex++)] = `Diff ${currencyPayment}`
    headers[this.toColumnLetter(columnIndex++)] = `Diff % ${currencyPayment}`

    // Additional information
    headers[this.toColumnLetter(columnIndex++)] = 'Effective From'
    headers[this.toColumnLetter(columnIndex++)] = 'Effective To'
    headers[this.toColumnLetter(columnIndex++)] = 'Lead Time Day' // Add Lead Time header
    headers[this.toColumnLetter(columnIndex++)] = 'CU-CR Reason'
    headers[this.toColumnLetter(columnIndex++)] = 'Original'
    headers[this.toColumnLetter(columnIndex++)] = 'Effective Model'
    headers[this.toColumnLetter(columnIndex++)] = 'Remark'

    return headers
  }

  // Helper function to map part cost object to Excel row
  private mapPartCostToExportRow(partCost: RequestPartCostDetail): { [key: string]: any } {
    const isBuyPart = this.isBuyPart()
    const currencyPart = this.getCurrencyName(
      this.partCostRequestData.requestPartCost.partMasterCurrencyId
    )
    const currencyPayment = this.getCurrencyName(
      this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
    )
    const unit = this.partCostRequestData.requestPartCost.unit || 'KG'

    // Tạo object để lưu trữ theo thứ tự chính xác của header
    const orderedRow: { [key: string]: any } = {}

    // Basic Information - match exact order as headers
    orderedRow['Is Parent Part'] = partCost.isParentPart ? 'Yes' : 'No'
    orderedRow['Is Upload SAP'] = partCost.isUploadSap ? 'Yes' : 'No'
    orderedRow['Is New Part'] = partCost.isNewPart ? 'Yes' : 'No'
    orderedRow['Base Part Code'] = partCost.basePartCode || ''
    orderedRow['New Part Code'] = partCost.newPartCode || ''
    orderedRow['Part Name/Spec'] = partCost.partNameSpec || ''    // Present values - ensure order matches headers
    if (isBuyPart) {
      orderedRow['Present Material Cost'] = partCost.presentMaterialCost ?? null
      orderedRow['Present Processing Cost'] = partCost.presentProcessingCost ?? null
      orderedRow['Present Other Cost'] = partCost.presentOtherCost ?? null
      orderedRow['Present Total Price'] = partCost.presentTotalPrice ?? null
    } else {
      orderedRow[`Present ${currencyPart}/KG`] = partCost.presentForeignCurrencyPerUnit ?? null
      orderedRow['Present Weight/pc'] = partCost.presentWeightPerUnit ?? null
      orderedRow['Present Other Cost'] = partCost.presentOtherCost ?? null
      orderedRow[`Present Total Price (${currencyPart}/pcs)`] = partCost.presentTotalPrice ?? null
    }

    // Present Other Cost - common for both Buy Part and Buy Material

    if (this.isDiffCurrency) {
      orderedRow['Present Exchange Rate'] = partCost.presentExchangeRate || 1
      orderedRow[`Present Total Price (${currencyPayment}/pcs)`] =
        partCost.presentTotalPriceConverted
    }

    orderedRow['Present MOQ'] = partCost.presentMoq || 0
    orderedRow['Present Incoterm'] = this.getIncotermName(partCost.presentIncotermId)
    orderedRow['Present Import Duty'] = partCost.presentImportDuty || 0
    orderedRow['Present Currency'] = currencyPayment

    // New values - ensure order matches headers
    if (isBuyPart) {
      orderedRow['New Material Cost'] = partCost.newMaterialCost ?? null
      orderedRow['New Processing Cost'] = partCost.newProcessingCost ?? null
      orderedRow['New Other Cost'] = partCost.newOtherCost ?? null
      orderedRow['New Total Price'] = partCost.newTotalPrice ?? null
    } else {
      orderedRow[`New ${currencyPart}/KG`] = partCost.newForeignCurrencyPerUnit ?? null
      orderedRow['New Weight/pc'] = partCost.newWeightPerUnit ?? null
      orderedRow['New Other Cost'] = partCost.newOtherCost ?? null
      orderedRow[`New Total Price (${currencyPart}/pcs)`] = partCost.newTotalPrice ?? null
    }

    // New Other Cost - common for both Buy Part and Buy Material

    if (this.isDiffCurrency) {
      orderedRow['New Exchange Rate'] = partCost.newExchangeRate || 1
      orderedRow[`New Total Price (${currencyPayment}/pcs)`] = partCost.newTotalPriceConverted || 0
    }

    orderedRow['New MOQ'] = partCost.newMoq || 0
    orderedRow['New Incoterm'] = this.getIncotermName(partCost.newIncotermId)
    orderedRow['New Import Duty'] = partCost.newImportDuty || 0
    orderedRow['New Currency'] = currencyPayment

    // Difference columns - ensure order matches headers
    if (this.isDiffCurrency) {
      orderedRow[`Diff ${currencyPart}`] = partCost.diffTotalPriceCurrencyPart || 0
      orderedRow[`Diff % ${currencyPart}`] = partCost.diffTotalPricePercentCurrencyPart || 0
    }
    orderedRow[`Diff ${currencyPayment}`] = partCost.diffTotalPricePayment || 0
    orderedRow[`Diff % ${currencyPayment}`] = partCost.diffTotalPricePercentPayment || 0

    // Additional Information - ensure order matches headers
    orderedRow['Effective From'] = partCost.effectiveDateFrom
      ? this.datePipe.transform(partCost.effectiveDateFrom, 'dd/MM/yyyy')
      : ''
    orderedRow['Effective To'] = partCost.effectiveDateTo
      ? this.datePipe.transform(partCost.effectiveDateTo, 'dd/MM/yyyy')
      : ''
    orderedRow['Lead Time Day'] = partCost.leadTimeDay
    orderedRow['CU-CR Reason'] = this.getCuCrReasonsText(partCost.requestPartCostDetailCrCuReasons)
    orderedRow['Original'] = this.getOriginalName(partCost.masterOriginalId)
    orderedRow['Effective Model'] = partCost.effectiveModel || ''
    orderedRow['Remark'] = partCost.remark || ''

    return orderedRow
  }

  // Also update the template download method to include the same styling
  downloadTemplateExcel(): void {
    // Create headers
    const headers = this.createExportHeaders()
    const headerArray = Object.values(headers)

    // Create empty data row
    const emptyRow = Array(headerArray.length).fill('')

    // Create data array with header and empty row
    const data = [headerArray, emptyRow]

    // Create workbook and worksheet
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.aoa_to_sheet(data)

    // Set column widths
    worksheet['!cols'] = Array(headerArray.length).fill({ wch: 15 })

    // Define custom styles
    const headerStyle = {
      fill: {
        fgColor: { rgb: 'DDEBF7' },
        patternType: 'solid'
      },
      font: { bold: true, color: { rgb: '000000' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'thin', color: { rgb: '000000' } },
        left: { style: 'thin', color: { rgb: '000000' } },
        right: { style: 'thin', color: { rgb: '000000' } }
      }
    }

    // Apply header styles - this is critical for styling to work
    for (let C = 0; C < headerArray.length; C++) {
      const headerCell = XLSX.utils.encode_cell({ r: 0, c: C })
      if (!worksheet[headerCell]) continue

      // We need to ensure the cell has a proper structure for styling
      if (!worksheet[headerCell].s) worksheet[headerCell].s = {}
      Object.assign(worksheet[headerCell].s, headerStyle)
    }

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Template')

    // Use a direct browser-based approach for downloading
    const excelBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
      cellStyles: true
    })

    const fileData = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })

    // Create download link
    const url = URL.createObjectURL(fileData)
    const link = document.createElement('a')
    link.href = url
    link.download = 'PartCostTemplate.xlsx'

    // Trigger download
    document.body.appendChild(link)
    link.click()

    // Cleanup
    setTimeout(() => {
      document.body.removeChild(link)
      URL.revokeObjectURL(url)
    }, 100)
  }

  importFromExcel(event: any): void {
    this.isLoading = true
    const file: File = event.files[0];
    if (!file) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please select a file'
      });
      this.isLoading = false;
      return;
    }

    // Tạo FileParameter từ file được chọn
    const fileParameter: FileParameter = {
      data: file,
      fileName: file.name
    };

    this.partCostRequestData.requestPartCostDetails = []
    this.partCostService.importPartCostRequestData(this.partCostRequestData.requestBase.id, this.partCostRequestData.requestPartCost.partMasterCurrencyId, this.partCostRequestData.requestPartCost.paymentMasterCurrencyId, this.partCostRequestData.requestPartCost.masterTypeFormRequestPartCostId,
      this.partCostRequestData.requestPartCost.isSameLeadTime ?? false, this.partCostRequestData.requestPartCost.leadTimeDay ?? 0, fileParameter).subscribe({
      next: (data : PartCostRequestDataResponseData) => {
        if (data.isError == true) {
          this.messageService.add({
            severity: 'error',
            summary: 'Error',
            detail: 'Please select a file'
          });
          this.isLoading = false
        } else {
          this.partCostRequestData.requestPartCostDetails = data.result.requestPartCostDetails
          console.log("abc, data.result.requestPartCostDetails", data.result.requestPartCostDetails)
                    this.messageService.add({
            severity: 'success',
            summary: 'success',
            detail: 'success'
          });
          this.isLoading = false
        }
        this.clearFileSelection()
      }, error: (err: any) => {
              this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Please select a file'
      })
       this.clearFileSelection()
      }
    })
  }


  // Helper methods for import
  private parseNumber(value: any): number {
    if (value === undefined || value === null || value === '') return null
    const num = parseFloat(String(value).replace(/,/g, ''))
    return isNaN(num) ? 0 : num
  }

  editMode() {
    this.isEditMode = true
    if (this.partCostRequestData.requestPartCostDetails.length > 0) {
      for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
        const partCost = this.partCostRequestData.requestPartCostDetails[i]
        const editButton = document.querySelector('#btn-edit-data-' + partCost.id)
        if (editButton) {
          ;(editButton as HTMLElement).click()
        }
      }
    }
  }
  viewMode() {
    this.isEditMode = false
    if (this.partCostRequestData.requestPartCostDetails.length > 0) {
      for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
        const partCost = this.partCostRequestData.requestPartCostDetails[i]
        const editButton = document.querySelector('#btn-view-data-' + partCost.id)
        if (editButton) {
          ;(editButton as HTMLElement).click()
        }
      }
    }
  }

  private parseDate(dateValue: any): Date | null {
    if (!dateValue) return null

    // If dateValue is already a Date object
    if (dateValue instanceof Date && isValid(dateValue)) {
      return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate(), 0, 0, 0)
    }

    // If it's a number (Excel serial date)
    if (typeof dateValue === 'number') {
      const offsetInMilliseconds = new Date().getTimezoneOffset() * 60 * 1000
      const excelEpoch = new Date(Date.UTC(1900, 0, 0))
      const resultDate = new Date(excelEpoch.getTime() + (dateValue - 1) * 24 * 60 * 60 * 1000)
      // Remove time component and adjust for timezone
      return new Date(
        resultDate.getUTCFullYear(),
        resultDate.getUTCMonth(),
        resultDate.getUTCDate(),
        0,
        0,
        0
      )
    }

    // If it's a string, try parsing with different formats
    if (typeof dateValue === 'string') {
      const formats = [
        'dd/MM/yyyy',
        'MM/dd/yyyy',
        'yyyy-MM-dd',
        'dd-MM-yyyy',
        'yyyy/MM/dd',
        'dd.MM.yyyy'
      ]

      for (const format of formats) {
        const parsedDate = parse(dateValue.trim(), format, new Date())
        if (isValid(parsedDate)) {
          // Remove time component
          return new Date(
            parsedDate.getFullYear(),
            parsedDate.getMonth(),
            parsedDate.getDate(),
            0,
            0,
            0
          )
        }
      }

      // Try direct parsing as last resort
      const directParsed = new Date(dateValue)
      if (isValid(directParsed)) {
        return new Date(
          directParsed.getFullYear(),
          directParsed.getMonth(),
          directParsed.getDate(),
          0,
          0,
          0
        )
      }
    }

    console.warn(`Could not parse date value: ${dateValue}`)
    return null
  }

  private findIncotermIdByName(incotermName: string): number | undefined {
    if (!incotermName) return undefined
    for (const [id, incoterm] of this.masterIncotermsMap.entries()) {
      if (incoterm.name === incotermName) {
        return id
      }
    }
    return undefined
  }

  private findOriginalIdByName(originalName: string): number | undefined {
    if (!originalName) return undefined
    for (const [id, original] of this.masterOriginalsMap.entries()) {
      if (original.name === originalName) {
        return id
      }
    }
    return undefined
  }

  // Improved implementation of CU-CR reasons mapping
  mapCuCrReasons(partCost: any, reasonsString: string): void {
    if (!reasonsString) {
      partCost.requestPartCostDetailCrCuReasons = []
      return
    }

    // Split by comma, trim whitespace, and filter out empty entries
    const reasonNames = reasonsString
      .split(',')
      .map((r) => r.trim())
      .filter((r) => r.length > 0)

    // Create a Set to avoid duplicates
    const matchedReasons = new Set()

    // First try exact matches (case-insensitive)
    for (const reasonName of reasonNames) {
      const reasonLower = reasonName.toLowerCase().trim()
      for (const [id, reason] of this.masterCrCuReasonsMap.entries()) {
        if (reason.masterDefinition.toLowerCase() === reasonLower) {
          matchedReasons.add(reason)
          break // Found exact match, no need to continue for this reason
        }
      }
    }

    // // If no exact matches were found, try partial matches
    // if (matchedReasons.size === 0) {
    //   for (const reasonName of reasonNames) {
    //     const reasonLower = reasonName.toLowerCase()
    //     for (const [id, reason] of this.masterCrCuReasonsMap.entries()) {
    //       const masterDefLower = reason.masterDefinition.toLowerCase()
    //       if (masterDefLower.includes(reasonLower) || reasonLower.includes(masterDefLower)) {
    //         matchedReasons.add(reason)
    //       }
    //     }
    //   }
    // }

    // Set the matched reasons
    partCost.requestPartCostDetailCrCuReasons = Array.from(matchedReasons)

    // If still no matches, add a message to help debugging
    if (partCost.requestPartCostDetailCrCuReasons.length === 0 && reasonNames.length > 0) {
      console.warn(`Could not match any CU-CR reasons for: ${reasonsString}`)
    }
  }

  private createNewEmptyPartCost(): any {
    const partCost = this.helperService.createNew(RequestPartCostDetail)
    partCost.id = this.helperService.generateUUID()
    partCost.requestBaseId = this.partCostRequestData.requestBase.id
    partCost.requestPartCostId = this.partCostRequestData.requestPartCost.id
    partCost.requestPartCostDetailCrCuReasons = []
    return partCost
  }

  // Helper method to ensure consistent toColumnLetter implementation
  toColumnLetter(columnNumber: number): string {
    let columnLetter = ''
    while (columnNumber >= 0) {
      const remainder = columnNumber % 26
      columnLetter = String.fromCharCode(65 + remainder) + columnLetter
      columnNumber = Math.floor(columnNumber / 26) - 1
      if (columnNumber < 0) break
    }
    return columnLetter
  }

  // Helper methods for Excel export/import
  private getIncotermName(incotermId: number | undefined): string {
    if (!incotermId) return ''
    const incoterm = this.masterIncotermsMap.get(incotermId)
    return incoterm ? incoterm.name : ''
  }

  private getOriginalName(originalId: number | undefined): string {
    if (!originalId) return ''
    const original = this.masterOriginalsMap.get(originalId)
    return original ? original.name : ''
  }

  private getCuCrReasonsText(reasons: MasterCrCuReason[]): string {
    if (!reasons || reasons.length === 0) return ''
    return reasons.map((r) => r.masterDefinition).join(', ')
  }

  // Add missing helper method to get currency name
  private getCurrencyName(currencyId: number): string {
    if (!currencyId) return ''
    const currency = this.masterCurrenciesMap.get(currencyId)
    return currency ? currency.name : ''
  }

  // Fix method to clear Excel file selection
  clearFileSelection() {
    this.isLoading = false
    if (this.fileUpload) {
      this.fileUpload.clear()
    }
  }

  changePaymentCurrency() {
    this.currencyPayment = this.masterCurrenciesMap.has(
      this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
    )
      ? this.masterCurrenciesMap.get(
          this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
        )
      : null
  }

  changePartCurrency() {
    this.currencyPart = this.masterCurrenciesMap.has(
      this.partCostRequestData.requestPartCost.partMasterCurrencyId
    )
      ? this.masterCurrenciesMap.get(this.partCostRequestData.requestPartCost.partMasterCurrencyId)
      : null
  }
  getCurrencyPaymentNumberFormat(): string {
    let round = this.currencyPayment?.roundCost ?? 0
    if (this.isView) {
      if (this.currencyPayment == null) {
        this.currencyPayment = this.masterCurrenciesMap.has(
          this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
        )
          ? this.masterCurrenciesMap.get(
              this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
            )
          : null
        round = this.currencyPayment?.roundCost ?? 0
      }
    }
    return `1.0-${round}`
  }
  getCurrencyPaymentPercentNumberFormat(): string {
    let round = this.currencyPayment?.roundPercent ?? 0
    if (this.isView) {
      if (this.currencyPayment == null) {
        this.currencyPayment = this.masterCurrenciesMap.has(
          this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
        )
          ? this.masterCurrenciesMap.get(
              this.partCostRequestData.requestPartCost.paymentMasterCurrencyId
            )
          : null
        round = this.currencyPayment?.roundPercent ?? 0
      }
    }
    return `1.0-${round}`
  }
  getCurrencyPartNumberFormat(): string {
    let round = this.currencyPart?.roundCost ?? 0
    if (this.isView) {
      if (this.currencyPart == null) {
        this.currencyPart = this.masterCurrenciesMap.has(
          this.partCostRequestData.requestPartCost.partMasterCurrencyId
        )
          ? this.masterCurrenciesMap.get(
              this.partCostRequestData.requestPartCost.partMasterCurrencyId
            )
          : null
        round = this.currencyPart?.roundCost ?? 0
      }
    }
    return `1.0-${round}`
  }
  getCurrencyPartPercentNumberFormat(): string {
    let round = this.currencyPart?.roundPercent ?? 0
    if (this.isView) {
      if (this.currencyPart == null) {
        this.currencyPart = this.masterCurrenciesMap.has(
          this.partCostRequestData.requestPartCost.partMasterCurrencyId
        )
          ? this.masterCurrenciesMap.get(
              this.partCostRequestData.requestPartCost.partMasterCurrencyId
            )
          : null
        round = this.currencyPart?.roundPercent ?? 0
      }
    }
    return `1.0-${round}`
  }

  changeIsParentPart(partCost: RequestPartCostDetail) {
    if (partCost.isParentPart == true) {
      partCost.isUploadSap = true
    }
  }

  changeEffectiveDate(partCost: RequestPartCostDetail) {
    if (
      partCost.effectiveDateFrom &&
      this.dateHelperService.isBefore(partCost.effectiveDateFrom, new Date())
    ) {
      this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId =
        MasterTypeRequestPartCostEnum.BackDate
      this.messageService.add({
        severity: 'warn',
        summary: 'Warning',
        detail: 'Request part cost is back date!'
      })
    }
  }

deleteDraftRequest() {
  // Kiểm tra điều kiện có thể xóa request draft
  if (!this.canDeleteDraftRequest()) {
    this.messageService.add({
      severity: 'error',
      summary: 'Error',
      detail: 'Cannot delete this request. Please check the conditions.'
    })
    return
  }

  // Sử dụng confirm dialog của PrimeNG thay vì SweetAlert2
  this.confirmationService.confirm({
    message: `Are you sure you want to delete this draft request?
              <br><strong>Request No:</strong> ${this.partCostRequestData.requestBase.requestNo}
              <br><strong>Request Name:</strong> ${this.partCostRequestData.requestBase.shortDescription}
              <br><span style="color: red;"><strong>Warning:</strong> This action cannot be undone. All related data and files will be permanently deleted.</span>`,
    header: 'Confirm Delete Draft Request',
    icon: 'pi pi-exclamation-triangle',
    acceptButtonStyleClass: 'p-button-danger',
    rejectButtonStyleClass: 'p-button-secondary',
    acceptLabel: 'Delete Request',
    rejectLabel: 'Cancel',
    accept: () => {
      // Gọi API xóa request nếu người dùng đã xác nhận
      this.isLoading = true
      this.partCostService
        .deleteDraftRequest(this.partCostRequestData.requestBase.requestNo.toString())
        .subscribe({
          next: (response: ResponseData) => {
            this.isLoading = false
            if (!response.isError) {
              // Xóa thành công
              this.messageService.add({
                severity: 'success',
                summary: 'Deleted!',
                detail: 'Draft request has been deleted successfully.'
              })
              // Redirect về trang danh sách requests
              setTimeout(() => {
                this.router.navigate(['/request_part_cost'])
              }, 1000)
            } else {
              // Xóa thất bại
              this.messageService.add({
                severity: 'error',
                summary: 'Error!',
                detail: response.message || 'Cannot delete request. Please try again later.'
              })
            }
          },
          error: (error: any) => {
            this.isLoading = false
            this.messageService.add({
              severity: 'error',
              summary: 'Error!',
              detail: 'An error occurred while deleting the request. Please try again later.'
            })
            console.error('Error deleting draft request:', error)
          }
        })
    }
  })
}

// Method kiểm tra điều kiện có thể xóa request draft
  canDeleteDraftRequest(): boolean {
    // Kiểm tra xem request có ở trạng thái draft không
    if (this.partCostRequestData.requestBase.masterRequestStatusId !== StatusRequestEnum.Draft) {
      return false
    }

    // Kiểm tra xem người dùng hiện tại có phải là người tạo request không
    if (this.partCostRequestData.requestBase.requestBy !== this.authClientService.getEmployeeCodeAuthed()) {
      return false
    }
     // Kiểm tra xem có approval summary nào không hoặc tất cả approval summary đều từ người dùng hiện tại
    if (this.partCostRequestData.requestApprovalSummaries &&
        this.partCostRequestData.requestApprovalSummaries.length > 0) {
      // Kiểm tra xem có approval summary nào từ người khác không
      const hasOtherUserApprovals = this.partCostRequestData.requestApprovalSummaries.some(
        summary => summary.createdBy !== this.authClientService.getEmployeeCodeAuthed()
      )
      if (hasOtherUserApprovals) {
        return false
      }
    }

    return true
  }

  calculateFirstSubmitDate(): void {
  // Khởi tạo firstSubmitDate với giá trị hiện tại
  let calculatedFirstSubmitDate = this.firstSubmitDate

  // Tìm tất cả requestApprovalSummaries có masterRequestActionId = ActionRequestEnum.Submit
  const submitApprovals = this.partCostRequestData.requestApprovalSummaries?.filter(
    summary => summary.masterRequestActionId === ActionRequestEnum.Submit
  )

  if (submitApprovals && submitApprovals.length > 0) {
    // Tìm Created_at nhỏ nhất trong các submit approvals
    const earliestSubmitApproval = submitApprovals.reduce((earliest, current) => {
      const currentDate = new Date(current.createdAt)
      const earliestDate = new Date(earliest.createdAt)
      return currentDate < earliestDate ? current : earliest
    })

    const earliestSubmitDate = new Date(earliestSubmitApproval.createdAt)

    // So sánh với firstSubmitDate hiện tại và chọn ngày nhỏ hơn
    if (earliestSubmitDate < calculatedFirstSubmitDate) {
      calculatedFirstSubmitDate = earliestSubmitDate
    }
  }

  // Cập nhật firstSubmitDate
  this.firstSubmitDate = calculatedFirstSubmitDate

  console.log('Calculated firstSubmitDate:', this.firstSubmitDate)
}

  leadTimeChange() {
    if (this.partCostRequestData.requestPartCost.leadTimeDay == null) {
      this.messageService.add({
        severity: 'warn',
        summary: 'Warning',
        detail: 'Remove lead time day'
      })
    } else {
      for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
        this.partCostRequestData.requestPartCostDetails[i].leadTimeDay =
          this.partCostRequestData.requestPartCost.leadTimeDay
      }
      this.messageService.add({
        severity: 'warn',
        summary: 'Warning',
        detail: 'Change lead time day'
      })
    }
  }

  getListRFIEmployee() {
    let arrayEmployee: string[] = []
    arrayEmployee.push(this.partCostRequestData.requestPartCost.createdBy)
    let historyRoutingsCanRfi = this.partCostRequestData.requestHistoryRoutings.filter((item) => {
      item.masterRoutingLevel < this.partCostRequestData.requestBase.masterRoutingLevel
    })
    for (let i = 0; i < historyRoutingsCanRfi.length; i++) {
      arrayEmployee.push(historyRoutingsCanRfi[i].employeeCode)
    }
    // console.log('arrayEmployee', arrayEmployee)
    return arrayEmployee
  }

  isSameLeadTimeChange() {
    if (this.partCostRequestData.requestPartCost.isSameLeadTime == true) {
      for (let i = 0; i < this.partCostRequestData.requestPartCostDetails.length; i++) {
        this.partCostRequestData.requestPartCostDetails[i].leadTime =
          this.partCostRequestData.requestPartCost.leadTime
      }
    }
  }

  colSpanPartCost() {
    let numberCol = 8
    if (this.isDiffCurrency) {
      numberCol += 2
    }
    return numberCol
  }

  /**
   * Phương thức xử lý validation khi người dùng thay đổi giá trị
   * Được gọi sau khi debounce để tránh tính toán quá nhiều
   */
  private performValidation(partCost: RequestPartCostDetail, field: string) {
    if (!partCost || !partCost.id) return

    // Đảm bảo đối tượng lỗi tồn tại cho part này
    if (!this.errors[partCost.id as string]) {
      this.errors[partCost.id as string] = {}
    }

    // // Kiểm tra các trường part code
    // if (field === 'basePartCode' || field === 'newPartCode') {
    //   this.validatePartCodeOnly(partCost, field)
    // }

    // Kiểm tra các trường liên quan đến ngày
    else if (field === 'effectiveDateFrom' || field === 'effectiveDateTo') {
      this.validateDateFields(partCost)
    }

    // Kiểm tra các trường liên quan đến cost
    else if (
      field.includes('Cost') ||
      field.includes('Price') ||
      field.includes('Weight') ||
      field.includes('Exchange') ||
      field.includes('Currency') ||
      field.includes('Moq') ||
      field.includes('ImportDuty')
    ) {
      // Chỉ xác thực trường liên quan đến cost
      this.validateCostField(partCost, field)
      // Sau khi xác thực xong, tính toán lại giá trị
    }

    // Kiểm tra các trường khác
    else {
      this.validateSingleField(partCost, field)
    }
  }

  /**
   * Phương thức xác thực ngày tháng
   */
  private validateDateFields(partCost: RequestPartCostDetail): boolean {
    let isValid = true

    // Kiểm tra ngày hiệu lực
    if (!partCost.effectiveDateFrom) {
      this.errors[partCost.id as string]['effectiveDateFrom'] = true
      isValid = false
    } else {
      this.errors[partCost.id as string]['effectiveDateFrom'] = false

      // Kiểm tra backdate
      if (
        this.dateHelperService.isBefore(partCost.effectiveDateFrom, new Date()) &&
        this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId !=
          MasterTypeRequestPartCostEnum.BackDate
      ) {
        this.partCostRequestData.requestPartCost.masterTypeRequestPartCostId =
          MasterTypeRequestPartCostEnum.BackDate
      }
    }

    if (!partCost.effectiveDateTo) {
      this.errors[partCost.id as string]['effectiveDateTo'] = true
      isValid = false
    } else {
      this.errors[partCost.id as string]['effectiveDateTo'] = false
    }

    return isValid
  }

  /**
   * Phương thức xác thực trường cost
   */
  private validateCostField(partCost: RequestPartCostDetail, field: string): boolean {
    let isValid = true

    switch (field) {
      case 'presentTotalPrice':
        if (
          (partCost.presentTotalPrice === null ||
           partCost.presentTotalPrice === undefined ||
           partCost.presentTotalPrice < 0) &&
          partCost.isNewPart == false
        ) {
          this.errors[partCost.id as string]['presentTotalPrice'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['presentTotalPrice'] = false
        }
        break

      case 'newTotalPrice':
        if (partCost.newTotalPrice === null ||
            partCost.newTotalPrice === undefined ||
            partCost.newTotalPrice < 0) {
          this.errors[partCost.id as string]['newTotalPrice'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['newTotalPrice'] = false
        }
        break

      case 'presentExchangeRate':
        if (this.isDiffCurrency && partCost.isParentPart) {
          if (
            !partCost.presentExchangeRate ||
            partCost.presentExchangeRate <= 0 ||
            partCost.presentExchangeRate == null
          ) {
            this.errors[partCost.id as string]['presentExchangeRate'] = true
            isValid = false
          } else {
            this.errors[partCost.id as string]['presentExchangeRate'] = false
          }
        }
        break

      case 'newExchangeRate':
        if (this.isDiffCurrency && partCost.isParentPart) {
          if (
            !partCost.newExchangeRate ||
            partCost.newExchangeRate <= 0 ||
            partCost.newExchangeRate == null
          ) {
            this.errors[partCost.id as string]['newExchangeRate'] = true
            isValid = false
          } else {
            this.errors[partCost.id as string]['newExchangeRate'] = false
          }
        }
        break

      case 'presentMoq':
        if (
          (partCost.presentMoq === undefined || partCost.presentMoq === null) &&
          partCost.isNewPart == false &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['presentMoq'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['presentMoq'] = false
        }
        break

      case 'newMoq':
        if ((partCost.newMoq === undefined || partCost.newMoq === null) && partCost.isParentPart) {
          this.errors[partCost.id as string]['newMoq'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['newMoq'] = false
        }
        break

      // Thêm các trường khác nếu cần
    }

    return isValid
  }

  /**
   * Phương thức xác thực một trường duy nhất
   */
  private validateSingleField(partCost: RequestPartCostDetail, field: string): boolean {
    let isValid = true

    switch (field) {
      case 'partNameSpec':
        if (!partCost.partNameSpec || partCost.partNameSpec.trim() === '') {
          this.errors[partCost.id as string]['partNameSpec'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['partNameSpec'] = false
        }
        break

      case 'masterOriginalId':
        if (!partCost.masterOriginalId && partCost.isParentPart) {
          this.errors[partCost.id as string]['masterOriginalId'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['masterOriginalId'] = false
        }
        break

      case 'presentIncotermId':
        if (
          (partCost.presentIncotermId === undefined || partCost.presentIncotermId === null) &&
          partCost.isNewPart == false &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['presentIncotermId'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['presentIncotermId'] = false
        }
        break

      case 'newIncotermId':
        if (
          (partCost.newIncotermId === undefined || partCost.newIncotermId === null) &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['newIncotermId'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['newIncotermId'] = false
        }
        break

      case 'presentImportDuty':
        if (
          (partCost.presentImportDuty === undefined || partCost.presentImportDuty === null) &&
          partCost.isNewPart == false &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['presentImportDuty'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['presentImportDuty'] = false
        }
        break

      case 'newImportDuty':
        if (
          (partCost.newImportDuty === undefined || partCost.newImportDuty === null) &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['newImportDuty'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['newImportDuty'] = false
        }
        break

      case 'requestPartCostDetailCrCuReasons':
        if (
          partCost.requestPartCostDetailCrCuReasons.length === 0 &&
          partCost.isNewPart == false &&
          partCost.isParentPart
        ) {
          this.errors[partCost.id as string]['requestPartCostDetailCrCuReasons'] = true
          isValid = false
        } else {
          this.errors[partCost.id as string]['requestPartCostDetailCrCuReasons'] = false
        }
        break
    }

    return isValid
  }

  getRoutings() {
    if (this.isShowCreate) {
      return this.partCostRequestData.masterRoutings
    } else return this.partCostRequestData.requestHistoryRoutings
  }

  get isShowCreate(): boolean {
    if (this.isView == false) return true
    if (
      this.isView == true &&
      this.isRequestDraft() &&
      this.partCostRequestData.requestApprovalSummaries.length == 0
    ) {
      return true
    }
    return false
  }

  get isShowReSubmit(): boolean {
    if (
      this.isView == true &&
      this.isRequestDraft() &&
      this.partCostRequestData.requestApprovalSummaries.length > 0
    ) {
      return true
    } else return false
  }

  masterUnit() {
    this.masterService.masterUnit().subscribe({
      next: (data: MasterUnit[]) => {
        this.masterUnits = data
        this.masterUnitsFilter = data.filter((x) => x.isActive == true)
        this.masterUnitsMap = new Map(
          data.map((d: MasterUnit) => {
            return [d.idNo, d]
          })
        )
        console.log('Master masterUnit:', this.masterUnit)
      },
      error: (error) => {
        console.error('Error fetching master masterUnit:', error)
      }
    })
  }

  ReCalcPartCost()
  {
     this.isLoading = true
    this.partCostService.reCalcPartCost(this.partCostRequestData).subscribe({
      next: (data : PartCostRequestDataResponseData) => {
        if (data.isError == true) {
          this.messageService.add({
            severity: 'error',
            summary: 'Error',
          });
          this.isLoading = false
        } else {
          this.partCostRequestData.requestPartCostDetails = data.result.requestPartCostDetails
          console.log("abc, data.result.requestPartCostDetails", data.result.requestPartCostDetails)
                    this.messageService.add({
            severity: 'success',
            summary: 'success',
            detail: 'success'
          });
          this.isLoading = false
        }
        this.clearFileSelection()
      }, error: (err: any) => {
              this.messageService.add({
        severity: 'error',
        summary: 'Error',
      })
       this.clearFileSelection()
      }
    })
  }

  // Thay thế hàm deleteRequestFile hiện tại bằng hàm mới
  deleteRequestFile(doc: RequestBaseSupportDocument) {
    // Hiển thị thông báo xác nhận với thông tin file
    Swal.fire({
      title: 'Xác nhận xóa file',
      html: `
      <div class="text-left">
        <p><strong>Tên file:</strong> ${doc.fileName}</p>
      </div>
    `,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      cancelButtonColor: '#3085d6',
      confirmButtonText: 'Xóa file',
      cancelButtonText: 'Hủy',
      focusCancel: true
    }).then((result) => {
      if (result.isConfirmed) {
        // Gọi API xóa file nếu người dùng đã xác nhận
        this.isLoading = true
        this.partCostService
          .deleteRequestFile(this.partCostRequestData.requestBase.requestNo.toString(), doc.id)
          .subscribe({
            next: (response: ResponseData) => {
              this.isLoading = false
              if (!response.isError) {
                // Xóa thành công
                Swal.fire('Đã xóa!', 'File đã được xóa thành công.', 'success')

                // Cập nhật lại danh sách file
                this.partCostRequestData.requestBaseSupportDocuments =
                  this.partCostRequestData.requestBaseSupportDocuments.filter(
                    (f) => f.id !== doc.id
                  )
                this.partCostRequestData.requestBaseQuotationDocuments =
                  this.partCostRequestData.requestBaseQuotationDocuments.filter(
                    (f) => f.id !== doc.id
                  )
              } else {
                // Xóa thất bại
                Swal.fire(
                  'Lỗi!',
                  response.message || 'Không thể xóa file. Vui lòng thử lại sau.',
                  'error'
                )
              }
            },
            error: (error: any) => {
              this.isLoading = false
              Swal.fire('Lỗi!', 'Đã xảy ra lỗi khi xóa file. Vui lòng thử lại sau.', 'error')
              console.error('Error deleting file:', error)
            }
          })
      }
    })  }

  // Dialog properties
  showEditPartCostDialog = false;
  editingPartCost: any = null;
  editingRowIndex = -1;
  originalPartCost: any = null; // Backup for cancel

  // Group editing properties
  editingPartCostGroup: any[] = [];
  originalPartCostGroup: any[] = [];
  currentGroupParentIndex = -1;

  getNavigationInfo(): string {
    if (!this.editingPartCostGroup || this.editingPartCostGroup.length === 0) {
      return 'No parts in group';
    }

    const parentPart = this.editingPartCostGroup.find(item => item.isParentPart);
    if (parentPart) {
      const parentCode = parentPart.basePartCode || parentPart.newPartCode || 'Unknown';
      return `Editing group: ${parentCode} (${this.editingPartCostGroup.length} parts)`;
    }

    return `Editing ${this.editingPartCostGroup.length} parts`;
  }

  onCloseEditDialog() {
    // Same as cancel for now
    this.cancelGroupEdit();
  }

  // Group editing methods
  editPartCostWithGroup(partCost: any, rowIndex: number) {
    // Find the parent and all children in the group
    const group = this.findPartCostGroup(partCost, rowIndex);

    // Add originalIndex property to each item in the group
    const allParts = this.partCostRequestData.requestPartCostDetails;
    this.editingPartCostGroup = group.map(item => {
      const originalIndex = allParts.findIndex(p =>
        p.basePartCode === item.basePartCode &&
        p.isParentPart === item.isParentPart
      );
      return { ...item, originalIndex };
    });

    this.originalPartCostGroup = group.map(item => ({ ...item }));
    this.currentGroupParentIndex = this.findParentIndexInGroup(group);

    this.showEditPartCostDialog = true;
  }

  private findPartCostGroup(targetPartCost: any, targetIndex: number): any[] {
    const allParts = this.partCostRequestData.requestPartCostDetails;

    // If target is a parent, find all its children
    if (targetPartCost.isParentPart) {
      const group = [targetPartCost];

      // Find all children after this parent
      for (let i = targetIndex + 1; i < allParts.length; i++) {
        const part = allParts[i];
        if (part.isParentPart) {
          break; // Found next parent, stop
        }
        group.push(part);
      }

      return group;
    } else {
      // If target is a child, find its parent and all siblings
      let parentIndex = -1;

      // Look backwards to find the parent
      for (let i = targetIndex - 1; i >= 0; i--) {
        if (allParts[i].isParentPart) {
          parentIndex = i;
          break;
        }
      }

      if (parentIndex === -1) {
        // No parent found, treat as standalone
        return [targetPartCost];
      }

      const group = [allParts[parentIndex]]; // Add parent

      // Add all children
      for (let i = parentIndex + 1; i < allParts.length; i++) {
        const part = allParts[i];
        if (part.isParentPart) {
          break; // Found next parent, stop
        }
        group.push(part);
      }

      return group;
    }
  }

  private findParentIndexInGroup(group: any[]): number {
    return group.findIndex(item => item.isParentPart);
  }

  getDialogTitle(): string {
    if (!this.editingPartCostGroup || this.editingPartCostGroup.length === 0) {
      return 'No parts selected';
    }

    const parentPart = this.editingPartCostGroup.find(item => item.isParentPart);
    if (parentPart) {
      return `${parentPart.basePartCode || 'Unknown'} Group`;
    }

    return 'Part Cost Group';
  }

  getGroupSummary(): string {
    if (!this.editingPartCostGroup || this.editingPartCostGroup.length === 0) {
      return 'No parts in group';
    }

    const parentCount = this.editingPartCostGroup.filter(item => item.isParentPart).length;
    const childCount = this.editingPartCostGroup.length - parentCount;

    return `Group contains ${parentCount} parent part(s) and ${childCount} child part(s)`;
  }

  getParentChildCount(): string {
    if (!this.editingPartCostGroup || this.editingPartCostGroup.length === 0) {
      return '0 parts';
    }

    const parentCount = this.editingPartCostGroup.filter(item => item.isParentPart).length;
    const childCount = this.editingPartCostGroup.length - parentCount;

    return `${parentCount} parent, ${childCount} children`;
  }

  onEditGroupFormChange() {
    // Trigger validation when group form changes
    this.validateRows();
  }

  editPreviousGroup() {
    // Find the previous parent part in the main list
    const currentParent = this.editingPartCostGroup.find(item => item.isParentPart);
    if (!currentParent) return;

    const allParts = this.partCostRequestData.requestPartCostDetails;
    const currentParentIndex = allParts.findIndex(item =>
      item.basePartCode === currentParent.basePartCode &&
      item.isParentPart
    );

    // Look backwards for the previous parent
    for (let i = currentParentIndex - 1; i >= 0; i--) {
      if (allParts[i].isParentPart) {
        this.saveGroupEdit();
        this.editPartCostWithGroup(allParts[i], i);
        return;
      }
    }
  }

  editNextGroup() {
    // Find the next parent part in the main list
    const currentParent = this.editingPartCostGroup.find(item => item.isParentPart);
    if (!currentParent) return;

    const allParts = this.partCostRequestData.requestPartCostDetails;
    const currentParentIndex = allParts.findIndex(item =>
      item.basePartCode === currentParent.basePartCode &&
      item.isParentPart
    );

    // Skip current group and look for next parent
    for (let i = currentParentIndex + this.editingPartCostGroup.length; i < allParts.length; i++) {
      if (allParts[i].isParentPart) {
        this.saveGroupEdit();
        this.editPartCostWithGroup(allParts[i], i);
        return;
      }
    }
  }

  canEditPreviousGroup(): boolean {
    const currentParent = this.editingPartCostGroup.find(item => item.isParentPart);
    if (!currentParent) return false;

    const allParts = this.partCostRequestData.requestPartCostDetails;
    const currentParentIndex = allParts.findIndex(item =>
      item.basePartCode === currentParent.basePartCode &&
      item.isParentPart
    );

    // Check if there's a parent before current one
    for (let i = currentParentIndex - 1; i >= 0; i--) {
      if (allParts[i].isParentPart) {
        return true;
      }
    }

    return false;
  }

  canEditNextGroup(): boolean {
    const currentParent = this.editingPartCostGroup.find(item => item.isParentPart);
    if (!currentParent) return false;

    const allParts = this.partCostRequestData.requestPartCostDetails;
    const currentParentIndex = allParts.findIndex(item =>
      item.basePartCode === currentParent.basePartCode &&
      item.isParentPart
    );

    // Check if there's a parent after current group
    for (let i = currentParentIndex + this.editingPartCostGroup.length; i < allParts.length; i++) {
      if (allParts[i].isParentPart) {
        return true;
      }
    }

    return false;
  }

  saveGroupEdit() {
    if (this.isEditGroupFormValid()) {
      // Update the main array with changes from editing group
      const allParts = this.partCostRequestData.requestPartCostDetails;

      this.editingPartCostGroup.forEach(editedPart => {
        const index = allParts.findIndex(part =>
          part.basePartCode === editedPart.basePartCode &&
          part.isParentPart === editedPart.isParentPart
        );

        if (index !== -1) {
          allParts[index] = { ...editedPart };
        }
      });

      // Update virtual scroll data
      this.updateVirtualScrollData();

      // Recalculate costs and validation
      this.ReCalcPartCost();
      this.validateRows();

      this.showEditPartCostDialog = false;
      this.editingPartCostGroup = [];
      this.originalPartCostGroup = [];
      this.currentGroupParentIndex = -1;

      this.messageService.add({
        severity: 'success',
        summary: 'Success',
        detail: 'Part cost group updated successfully'
      });
    }
  }

  cancelGroupEdit() {
    // Restore original data if needed
    this.showEditPartCostDialog = false;
    this.editingPartCostGroup = [];
    this.originalPartCostGroup = [];
    this.currentGroupParentIndex = -1;
  }

  isEditGroupFormValid(): boolean {
    if (!this.editingPartCostGroup || this.editingPartCostGroup.length === 0) {
      return false;
    }

    // Validate each part in the group
    for (const partCost of this.editingPartCostGroup) {
      const requiredFields = ['basePartCode', 'newPartCode', 'partNameSpec', 'presentTotalPrice'];

      for (const field of requiredFields) {
        if (!partCost[field] || partCost[field].toString().trim() === '') {
          return false;
        }
      }

      // Date validation
      if (!partCost.effectiveDateFrom || !partCost.effectiveDateTo) {
        return false;
      }
    }

    return true;
  }
  updateVirtualScrollData() {
    // Implement virtual scroll update if needed
    // This method updates the display array for virtual scrolling
    // For now, we'll just trigger a change detection
    this.validateRows();
  }

  	  @HostListener('document:keydown', ['$event'])
  handleKeyDown(event: KeyboardEvent): void {
    if (event.ctrlKey && event.key === 'p') {
      event.preventDefault();
      // this.directPrint();
    }
  }


  closePrintPreview(): void {
    this.showPrintPreview = false;
  }


  // New method for showing preview dialog (optional)
  showPrintPreviewDialog(): void {
    console.log('showPrintPreviewDialog called');
    console.log('partCostRequestData:', this.partCostRequestData);

    // Validate that we have data to print
    if (!this.partCostRequestData || !this.partCostRequestData.requestBase) {
      console.log('No data available for printing');
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'No data available for printing. Please load the request first.'
      });
      return;
    }

    // Check if there are part cost details to print
    if (!this.partCostRequestData.requestPartCostDetails ||
        this.partCostRequestData.requestPartCostDetails.length === 0) {
      console.log('No part cost details found');
      this.messageService.add({
        severity: 'warn',
        summary: 'Warning',
        detail: 'No part cost details found to print.'
      });
      return;
    }

    // Allow printing even if request is not submitted yet (for draft review)
    if (!this.partCostRequestData.requestBase.requestNo ||
        this.partCostRequestData.requestBase.requestNo === this.helperService.intMin) {
      console.log('Request not submitted yet - showing draft preview');
      this.messageService.add({
        severity: 'info',
        summary: 'Draft Preview',
        detail: 'Showing preview of draft request (not yet submitted).'
      });
    }

    console.log('Opening print preview...');
    this.showPrintPreview = true;
    console.log('showPrintPreview set to:', this.showPrintPreview);
  }


  // New method to create print window
  private createPrintWindow(): void {
    // Create temporary component instance for printing
    this.showPrintPreview = true;

    // Wait for component to render
    setTimeout(() => {
      // Get the print content
      const printContent = document.querySelector('app-part-cost-print');

      if (!printContent) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'Print content not found'
        });
        this.showPrintPreview = false;
        return;
      }

      // Create print window
      const printWindow = window.open('', '_blank', 'width=1200,height=800');

      if (!printWindow) {
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'Unable to open print window. Please check popup blocker settings.'
        });
        this.showPrintPreview = false;
        return;
      }      // Write content to print window with improved print styles
      printWindow.document.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <title>Part Cost Request - ${this.partCostRequestData.requestBase.requestNo}</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
              /* Page layout simulation */
              body {
                font-family: Arial, sans-serif;
                margin: 20px auto;
                padding: 0;
                background: #f5f5f5;
                font-size: 10px;
                line-height: 1.2;
              }

              /* Container to simulate A4 landscape page */
              .page-container {
                width: 297mm; /* A4 landscape width */
                min-height: 210mm; /* A4 landscape height */
                margin: 20px auto;
                padding: 0.5in 0.3in;
                background: white;
                box-shadow: 0 0 20px rgba(0,0,0,0.1);
                page-break-after: always;
              }

              /* Print buttons for preview */
              .print-buttons {
                position: fixed;
                top: 10px;
                right: 10px;
                z-index: 9999;
                background: white;
                padding: 10px;
                border-radius: 5px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.2);
              }

              .print-btn, .close-btn {
                padding: 8px 12px;
                margin: 0 3px;
                border: none;
                border-radius: 3px;
                cursor: pointer;
                font-size: 12px;
              }

              .print-btn {
                background: #007bff;
                color: white;
              }

              .close-btn {
                background: #6c757d;
                color: white;
              }

              /* Table styles to match print */
              .data-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 8px;
                table-layout: fixed;
                margin-bottom: 10px;
              }

              .data-table th,
              .data-table td {
                border: 1px solid black;
                padding: 2px 1px;
                font-size: 8px;
                line-height: 1.1;
                word-wrap: break-word;
                vertical-align: top;
              }

              .header-cell {
                background-color: #f8f9fa;
                color: black;
                font-weight: bold;
                text-align: center;
              }

              .data-cell.number {
                text-align: right;
                font-family: monospace;
              }

              .data-cell.text-center {
                text-align: center;
              }

              /* Header styles */
              .print-header {
                display: flex;
                justify-content: space-between;
                margin-bottom: 15px;
                border-bottom: 2px solid #333;
                padding-bottom: 10px;
              }

              .company-info h1 {
                margin: 0;
                font-size: 24px;
                color: #333;
              }

              .company-info h2 {
                margin: 5px 0 0 0;
                font-size: 18px;
                color: #666;
              }

              .request-info {
                text-align: right;
              }

              .info-row {
                margin: 5px 0;
              }

              .label {
                font-weight: bold;
                margin-right: 10px;
              }

              /* Section styles */
              .general-details {
                margin-bottom: 15px;
              }

              .general-details h3 {
                background: #f8f9fa;
                padding: 8px;
                margin: 0 0 10px 0;
                border-left: 4px solid #007bff;
                font-size: 12px;
              }

              .details-grid {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 10px;
              }

              .detail-item {
                padding: 5px 0;
              }

              /* Explanation content tables */
              .explanation-content table,
              .explanation-content th,
              .explanation-content td {
                border: 1px solid black !important;
                border-collapse: collapse !important;
                padding: 2px 4px !important;
                font-size: 8px !important;
              }

              /* Colors for positive/negative values */
              .positive {
                color: green;
              }

              .negative {
                color: red;
              }

              .bold {
                font-weight: bold;
              }

              /* Hide print controls in actual print */
              @media print {
                .print-buttons {
                  display: none !important;
                }

                body {
                  background: white;
                  margin: 0;
                }

                .page-container {
                  width: 100%;
                  min-height: auto;
                  margin: 0;
                  padding: 0.3in;
                  box-shadow: none;
                  page-break-after: auto;
                }

                @page {
                  size: A4 landscape;
                  margin: 0.5in 0.3in;
                }
              }

              /* Responsive adjustments */
              @media (max-width: 1200px) {
                .page-container {
                  width: 95%;
                  margin: 10px auto;
                }
              }
            </style>
          </head>
          <body>
            <div class="print-buttons">
              <button class="print-btn" onclick="window.print()">🖨️ Print</button>
              <button class="close-btn" onclick="window.close()">❌ Close</button>
            </div>

            <div class="page-container">
              ${printContent.innerHTML}
            </div>
          </body>
        </html>
      `);      printWindow.document.close();

      // Focus the new window for preview (don't auto-print)
      printWindow.focus();

      console.log('Print preview window opened successfully');

      // Hide the temporary preview in parent window
      this.showPrintPreview = false;
    }, 100);
  }
}
