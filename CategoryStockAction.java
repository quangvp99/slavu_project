package vnm.web.action.stock;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.struts2.ServletActionContext;

import com.opensymphony.xwork2.Action;
import com.viettel.core.entities.CycleCount;
import com.viettel.core.entities.CycleCountMapProduct;
import com.viettel.core.entities.Product;
import com.viettel.core.entities.ProductInfo;
import com.viettel.core.entities.ProductLot;
import com.viettel.core.entities.Shop;
import com.viettel.core.entities.Staff;
import com.viettel.core.entities.enumtype.ActiveType;
import com.viettel.core.entities.enumtype.CycleCountType;
import com.viettel.core.entities.enumtype.CycleType;
import com.viettel.core.entities.enumtype.KPaging;
import com.viettel.core.entities.enumtype.ProductType;
import com.viettel.core.entities.enumtype.StockObjectType;
import com.viettel.core.entities.enumtype.StockTotalVOFilter;
import com.viettel.core.entities.vo.ObjectVO;
import com.viettel.core.entities.vo.StockTotalVO;
import com.viettel.core.exceptions.BusinessException;

import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jxls.transformer.XLSTransformer;
import viettel.passport.client.ShopToken;
import vnm.web.action.general.AbstractAction;
import vnm.web.bean.CycleCountMapBean;
import vnm.web.constant.ConstantManager;
import vnm.web.enumtype.FileExtension;
import vnm.web.enumtype.ShopReportTemplate;
import vnm.web.helper.Configuration;
import vnm.web.log.LogUtility;
import vnm.web.utils.DateUtil;
import vnm.web.utils.ReportUtils;
import vnm.web.utils.StringUtil;
import vnm.web.utils.ValidateUtil;

public class CategoryStockAction extends AbstractAction {
	private static final long serialVersionUID = -2128120975032582873L;
	private Staff staff = null;
	private Shop shop = null;
	private String startDate;
	private Integer cycleType;
	private String shopCode;
	private Long shopId;
	private String cycleCountCode;
	private CycleCount cycleCount;
	private Integer sortBy;
	private Integer firstNumber;
	private List<ProductInfo> lstCategoryType;
	private List<ProductInfo> lstSubCategoryType;
	private List<ProductInfo> lstSubCategoryCat;
	private List<String> lstProduct;
	ObjectVO<CycleCountMapProduct> lstCCMapProduct;
	private String productCode;
	private String productName;
	private Long category;
	private Long sub_category;
	private Integer fromAmnt;
	private Integer toAmnt;
	private Long cycleMapProductId;
	private String cycleCountDes;
	private List<Integer> lstIsDelete;
	private List<Integer> lstStockCardNumber;
	private Integer lastNumber;
	private Integer status;
	private Integer noPaging;
	private String listProductEx;

	private HashMap<String, Object> parametersReport;

	private Long idParentCat;

	@Override
	public HashMap<String, Object> getParametersReport() {
		return this.parametersReport;
	}

	public void setParametersReport(HashMap<String, Object> parametersReport) {
		this.parametersReport = parametersReport;
	}

	public List<ProductInfo> getLstSubCategoryCat() {
		return this.lstSubCategoryCat;
	}

	public void setLstSubCategoryCat(List<ProductInfo> lstSubCategoryCat) {
		this.lstSubCategoryCat = lstSubCategoryCat;
	}

	public Long getIdParentCat() {
		return this.idParentCat;
	}

	public void setIdParentCat(Long idParentCat) {
		this.idParentCat = idParentCat;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	@Override
	public void prepare() throws Exception {
		super.prepare();
		try {
			if ((this.currentUser != null) && (this.currentUser.getUserName() != null)) {
				this.staff = this.staffMgr.getStaffByCode(this.currentUser.getUserName());
				if ((this.staff != null) && (this.staff.getShop() != null)) {
					this.shop = this.staff.getShop();
					this.shopId = this.shop.getId();
					this.shopCode = this.shop.getShopCode();
				}
			}
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		}
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	@Override
	public String execute() {
		this.resetToken(this.result);
		try {
			this.lstCategoryType = new ArrayList<>();
			ObjectVO<ProductInfo> tmp = this.productInfoMgr.getListProductInfoStock(null, null, null, null, ActiveType.RUNNING, ProductType.CAT);
			if (tmp != null) {
				this.lstCategoryType = tmp.getLstObject();
			}
			tmp = this.productInfoMgr.getListProductInfoStock(null, null, null, null, ActiveType.RUNNING, ProductType.SUB_CAT);
			this.lstSubCategoryType = new ArrayList<>();
			if (tmp != null) {
				this.lstSubCategoryType = tmp.getLstObject();
			}
			if ((this.currentUser != null) && (this.currentUser.getUserName() != null)) {
				this.staff = this.staffMgr.getStaffByCode(this.currentUser.getUserName());
				if ((this.staff != null) && (this.staff.getShop() != null)) {
					// lay danh sack kiem ke
					ObjectVO<CycleCount> lstCCTmp = this.cycleCountMgr.getListCycleCount(null, null, CycleCountType.ONGOING, null, this.staff.getShop().getId(), null, null, null, this.getStrListShopId());
					if ((lstCCTmp != null) && (lstCCTmp.getLstObject() != null) && (lstCCTmp.getLstObject().size() > 0)) {
						// lay danh kiem ke moi nhat theo start_date
						this.cycleCount = lstCCTmp.getLstObject().get(0);
						if (this.cycleCount.getCycleType() != null) {
							this.cycleType = this.cycleCount.getCycleType().getValue();
						}
						this.status = this.cycleCount.getStatus().getValue();
						// kiem tra kiem ke co san pham kiem ke hay chua, co thi khong lam gi het, ngc lai: cho them danh sach mat hang
						if (this.cycleCountMgr.checkIfAnyProductCounted(this.cycleCount.getId())) {
							this.status = -1;
						}

					} else {
						this.status = -1;
					}
				}
			}
		} catch (BusinessException e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		}
		return Action.SUCCESS;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	public String getListSubCat() {
		try {
			this.lstSubCategoryCat = new ArrayList<>();
			this.lstSubCategoryCat = this.productInfoMgr.getListSubCat(ActiveType.RUNNING, ProductType.SUB_CAT, this.idParentCat);
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		}
		return Action.SUCCESS;
	}

	public String search() {

		this.result.put("page", this.page);
		this.result.put("max", this.max);
		try {
			KPaging<CycleCountMapProduct> kPaging = new KPaging<>();
			kPaging.setPage(this.page - 1);
			kPaging.setPageSize(this.max);
			if (!StringUtil.isNullOrEmpty(this.cycleCountCode)) {
				this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, this.shopId);
				this.lstCCMapProduct = this.cycleCountMgr.getListCycleCountMapProductByCycleCountId(null, this.cycleCount.getId());
				if (this.lstCCMapProduct != null) {
					this.result.put("total", this.lstCCMapProduct.getkPaging().getTotalRows());
					this.result.put("rows", this.lstCCMapProduct.getLstObject());
				}
			}
			this.result.put(Action.ERROR, false);
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
			this.result.put(Action.ERROR, true);
			this.errMsg = Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error");
			this.result.put("errMsg", this.errMsg);
		}
		return AbstractAction.JSON;
	}

	/**
	 *
	 *
	 * @param
	 * @author tientv
	 */
	public String saveProduct() {
		this.resetToken(this.result);
		boolean error = true;
		try {
			//Phieu kiem ke khong ton tai
			if (StringUtil.isNullOrEmpty(this.cycleCountCode)) {
				this.result.put(Action.ERROR, error);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));
				return AbstractAction.JSON;
			}
			if (this.shop == null) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", ValidateUtil.getErrorMsg(ConstantManager.ERR_NOT_EXIST_DB, Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.countinginput.shop.is.null")));
				return AbstractAction.JSON;
			}
			this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, this.shopId);
			if (this.cycleCount == null) {
				this.result.put(Action.ERROR, error);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));
				return AbstractAction.JSON;
			}
			if (this.cycleCount.getStatus() != CycleCountType.ONGOING) {
				this.result.put(Action.ERROR, error);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.not.ongoing"));
				return AbstractAction.JSON;
			}
			Date sysDate = this.commonMgr.getSysDate();
			this.cycleCount.setFirstNumber(this.firstNumber);
			this.cycleCount.setShop(this.shop);
			this.cycleCount.setCycleType(CycleType.parseValue(this.cycleType));
			this.cycleCount.setUpdateDate(sysDate);
			this.cycleCount.setUpdateUser(this.staff.getStaffCode());
			Product product;
			List<CycleCountMapProduct> lstCountMapProducts = new ArrayList<>();
			if (this.lstProduct != null) {
				if (this.lstProduct.size() == 0) {
					this.result.put(Action.ERROR, error);
					this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.list.product"));
					return AbstractAction.JSON;
				}
				for (int i = 0; i < this.lstProduct.size(); ++i) {
					CycleCountMapProduct countMapProduct = new CycleCountMapProduct();
					countMapProduct.setCycleCount(this.cycleCount);
					countMapProduct.setCreateUser(this.staff.getStaffCode());
					countMapProduct.setStockCardNumber(this.cycleCount.getFirstNumber());
					product = this.productMgr.getProductByCode(this.lstProduct.get(i));
					if (this.lstIsDelete.get(i) == 1) {
						CycleCountMapProduct cycleCountMapProduct = this.cycleCountMgr.getCycleCountMapProduct(this.cycleCount.getId(), product.getId());
						if (cycleCountMapProduct != null) {
							this.cycleCountMgr.deleteCycleCountMapProduct(cycleCountMapProduct);
						}
						continue;
					}
					if (product == null) {
						continue;
					}
					if (!this.cycleCountMgr.checkIfCycleCountMapProductExist(this.cycleCount.getId(), this.lstProduct.get(i), null, null)) {
						countMapProduct.setProduct(product);
						countMapProduct.setStockCardNumber(this.lstStockCardNumber.get(i));
						lstCountMapProducts.add(countMapProduct);
					} else {
						CycleCountMapProduct temp = this.cycleCountMgr.getCycleCountMapProduct(this.cycleCount.getId(), product.getId());
						if (temp != null) {
							temp.setStockCardNumber(this.lstStockCardNumber.get(i));
							temp.setUpdateDate(sysDate);
							temp.setUpdateUser(this.staff.getStaffCode());
							lstCountMapProducts.add(temp);
						}
					}
				}
				this.cycleCountMgr.createListCycleCountMapProduct(this.cycleCount, lstCountMapProducts, false, this.getLogInfoVO());
				//	error=false;
			} else {
				this.cycleCountMgr.createListCycleCountMapProduct(this.cycleCount, null, true, this.getLogInfoVO());
				//	error=false;
			}
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
			this.result.put(Action.ERROR, error);
			this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error"));
		}
		return AbstractAction.JSON;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	public String searchDetail() {
		this.result.put("list", new ArrayList<CycleCountMapProduct>());
		try {
			if (!StringUtil.isNullOrEmpty(this.cycleCountCode)) {
				this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, this.shopId);
				if (this.cycleCount == null) {
					this.result.put(Action.ERROR, true);
					this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));
					return AbstractAction.JSON;
				}
				this.lstCCMapProduct = this.cycleCountMgr.getListCycleCountMapProductByCycleCountId(null, this.cycleCount.getId());
				if (this.lstCCMapProduct == null) {
					this.result.put(Action.ERROR, true);
					return AbstractAction.JSON;
				}
				if (this.lstCCMapProduct.getLstObject().size() > 0) {
					this.result.put("list", this.lstCCMapProduct.getLstObject());
				} else {
					this.result.put("list", new ArrayList<CycleCountMapProduct>());
				}
				this.result.put(Action.ERROR, false);
			}
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
			this.result.put(Action.ERROR, true);
			this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error"));
			return AbstractAction.JSON;
		}
		return AbstractAction.JSON;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	public String searchProduct() {
		if (this.noPaging == null) {
			try {
				this.result.put("page", this.page);
				this.result.put("max", this.max);
				if (this.currentUser == null) {
					this.result.put(Action.ERROR, true);
					this.result.put("errMsg", ValidateUtil.getErrorMsg(ConstantManager.ERR_NOT_PERMISSION, ""));
					return AbstractAction.JSON;
				}
				ShopToken shT = this.currentUser.getShopRoot();
				if (shT == null) {
					this.result.put(Action.ERROR, true);
					this.result.put("errMsg", ValidateUtil.getErrorMsg(ConstantManager.ERR_NOT_PERMISSION, ""));
					return AbstractAction.JSON;
				}
				if (!StringUtil.isNullOrEmpty(this.cycleCountCode)) {
					KPaging<StockTotalVO> kPaging = new KPaging<>();
					kPaging.setPage(this.page - 1);
					kPaging.setPageSize(this.max);
					this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, shT.getShopId());
					if ((this.cycleCount != null) && CycleCountType.ONGOING.equals(this.cycleCount.getStatus()) && this.cycleCount.getShop().getId().equals(shT.getShopId())) {
						List<Long> listProduct = new ArrayList<>();
						if (!StringUtil.isNullOrEmpty(this.listProductEx)) {
							for (String str : this.listProductEx.split(",")) {
								listProduct.add(Long.valueOf(str));
							}
						}
						StockTotalVOFilter filter = new StockTotalVOFilter();
						filter.setProductCode(this.productCode);
						filter.setProductName(this.productName);
						filter.setShopId(shT.getShopId());
						filter.setExceptProductId(listProduct);
						filter.setWarehouseId(this.cycleCount.getWarehouse().getId());
						filter.setCatId(this.category); // id nganh hang
						filter.setSubCatId(this.sub_category);//id nganh hang con phu
						filter.setFromQuantity(this.fromAmnt);
						filter.setToQuantity(this.toAmnt);
						ObjectVO<StockTotalVO> stockTotalVO = this.stockMgr.getListProductForCycleCount(kPaging, this.cycleCount.getId(), filter);
						if (stockTotalVO != null) {
							this.result.put("total", stockTotalVO.getkPaging().getTotalRows());
							this.result.put("rows", stockTotalVO.getLstObject());
						}
					}
				}
			} catch (Exception e) {
				LogUtility.logError(e, e.getMessage(), this.getClass(), null);
			}
		} else if (this.noPaging.intValue() == 1) {
			try {
				if (this.currentUser == null) {
					this.result.put(Action.ERROR, true);
					this.result.put("errMsg", ValidateUtil.getErrorMsg(ConstantManager.ERR_NOT_PERMISSION, ""));
					return AbstractAction.JSON;
				}
				ShopToken shT = this.currentUser.getShopRoot();
				if (shT == null) {
					this.result.put(Action.ERROR, true);
					this.result.put("errMsg", ValidateUtil.getErrorMsg(ConstantManager.ERR_NOT_PERMISSION, ""));
					return AbstractAction.JSON;
				}
				if (!StringUtil.isNullOrEmpty(this.cycleCountCode)) {
					this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, shT.getShopId());
					if ((this.cycleCount != null) && CycleCountType.ONGOING.equals(this.cycleCount.getStatus()) && this.cycleCount.getShop().getId().equals(shT.getShopId())) {
						StockTotalVOFilter filter = new StockTotalVOFilter();
						filter.setProductCode(this.productCode);
						filter.setProductName(this.productName);
						filter.setShopId(shT.getShopId());
						filter.setWarehouseId(this.cycleCount.getWarehouse().getId());
						ObjectVO<StockTotalVO> stockTotalVO = this.stockMgr.getListProductForCycleCountCategoryStock(null, this.cycleCount.getId(), filter);
						if ((stockTotalVO != null) && (stockTotalVO.getLstObject().size() > 0)) {
							this.result.put("list", stockTotalVO.getLstObject());
						} else {
							this.result.put("list", new ArrayList<StockTotalVO>());
						}
						this.result.put(Action.ERROR, false);
					}
				}
			} catch (BusinessException e) {
				LogUtility.logError(e, e.getMessage(), this.getClass(), null);
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error"));
				return AbstractAction.JSON;
			}
		}
		return AbstractAction.JSON;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param
	 * @author tientv
	 */
	public String deleteCycleCountDetail() {
		this.resetToken(this.result);
		boolean error = true;
		String errMsg = "";
		try {
			CycleCountMapProduct cycleCountMapProduct = this.cycleCountMgr.getCycleCountMapProductById(this.cycleMapProductId);
			if (cycleCountMapProduct != null) {
				if (cycleCountMapProduct.getCycleCount().getStatus() == CycleCountType.ONGOING) {
					this.cycleCountMgr.deleteCycleCountMapProduct(cycleCountMapProduct);
					error = false;
				} else {
					errMsg = Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.not.ongoing");
				}
			}

		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		}
		this.result.put(Action.ERROR, error);
		if (error && StringUtil.isNullOrEmpty(errMsg)) {
			errMsg = Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error");
		}
		this.result.put("errMsg", errMsg);
		return AbstractAction.JSON;
	}

	/**
	 * Export PDF with Jasperreport
	 *
	 * @param
	 * @author hunglm16
	 * @since March 06, 2014
	 * @description Fix bug
	 */
	public String exportPdf() {
		this.parametersReport = new HashMap<>();
		this.formatType = FileExtension.PDF.getName();
		try {
			if (StringUtil.isNullOrEmpty(this.cycleCountCode)) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));

				return AbstractAction.JSON;
			}
			this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, this.shopId);
			if (this.cycleCount == null) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));
				return AbstractAction.JSON;
			}
			this.lstCCMapProduct = this.cycleCountMgr.getListCycleCountMapProduct(null, this.cycleCount.getId(), null, null, null, null, null, null, false);
			if (this.lstCCMapProduct == null) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.print"));
				return AbstractAction.JSON;
			}
			if (this.lstCCMapProduct.getLstObject().size() == 0) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.print"));
				return AbstractAction.JSON;
			}
			String printDate = DateUtil.toDateString(new Date(), DateUtil.DATE_FORMAT_NOW);
			this.parametersReport.put("printDate", printDate);
			if (this.shopId != null) {
				Shop shop = this.shopMgr.getShopById(this.shopId);
				if (shop != null) {
					this.parametersReport.put("shopName", shop.getShopName());
					if (!StringUtil.isNullOrEmpty(shop.getAddress())) {
						this.parametersReport.put("address", shop.getAddress());
					} else {
						this.parametersReport.put("address", "");
					}
				}
			}
			this.parametersReport.put("logoPath", ReportUtils.getVinamilkLogoRealPath(this.request));
			this.parametersReport.put("cycleCountCode", this.cycleCount.getCycleCountCode());
			List<CycleCountMapBean> lst = new ArrayList<>();
			for (CycleCountMapProduct ccMap : this.lstCCMapProduct.getLstObject()) {
				if (ccMap.getProduct().getCheckLot() != 1) {
					CycleCountMapBean a = new CycleCountMapBean();
					a.setStockCardNumber(ccMap.getStockCardNumber());
					a.setProductCode(ccMap.getProduct().getProductCode());
					a.setProductName(ccMap.getProduct().getProductName());
					a.setProductLot("");
					//					a.setQuantityCounted(ccMap.getQuantityCounted());
					//					a.setDescription(ccMap.getCycleCount().getDescription());
					lst.add(a);
				} else {//loctt - Oct14, 2013
					List<ProductLot> productLot = this.productMgr.getProductLotByProductAndOwner(ccMap.getProduct().getId(), this.shopId, StockObjectType.SHOP);
					for (ProductLot pl : productLot) {
						CycleCountMapBean a1 = new CycleCountMapBean();
						a1.setStockCardNumber(ccMap.getStockCardNumber());
						a1.setProductCode(ccMap.getProduct().getProductCode());
						a1.setProductName(ccMap.getProduct().getProductName());
						a1.setProductLot(pl.getLot());
						lst.add(a1);
					}
				}
			}
			//if (lst != null ) {
			lst.add(0, new CycleCountMapBean());
			//}
			this.session.setAttribute(ConstantManager.SESSION_REPORT_DATA, lst);
			this.session.setAttribute(ConstantManager.SESSION_REPORT_PARAM, this.parametersReport);
			FileExtension ext = FileExtension.parseValue("PDF");

			JRDataSource dataSource = new JRBeanCollectionDataSource(lst);
			String outputPath = ReportUtils.exportFromFormat(ext, this.parametersReport, dataSource, ShopReportTemplate.STOCK_CATEGORY_INFOR_EXPORT_INVENTORY);
			this.result.put(Action.ERROR, false);
			this.result.put(AbstractAction.LIST, outputPath);
			return AbstractAction.JSON;
		} catch (Exception e) {
			this.result.put(Action.ERROR, true);
			this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error"));
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		}
		return AbstractAction.JSON;
	}

	/**
	 *
	 *
	 * @param
	 * @author tientv
	 */
	public String exportExcel() {
		InputStream inputStream = null;
		OutputStream os = null;
		Workbook resultWorkbook = null;
		try {
			if (StringUtil.isNullOrEmpty(this.cycleCountCode)) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));

				return AbstractAction.JSON;
			}
			this.cycleCount = this.cycleCountMgr.getCycleCountByCodeAndShop(this.cycleCountCode, this.shopId);
			if (this.cycleCount == null) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.cyclecount.code.null"));
				return AbstractAction.JSON;
			}
			this.lstCCMapProduct = this.cycleCountMgr.getListCycleCountMapProduct(null, this.cycleCount.getId(), null, null, null, null, null, null, false);
			if (this.lstCCMapProduct == null) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.print"));
				return AbstractAction.JSON;
			}
			if (this.lstCCMapProduct.getLstObject().size() == 0) {
				this.result.put(Action.ERROR, true);
				this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "stock.category.print"));
				return AbstractAction.JSON;
			}
			String printDate = DateUtil.toDateString(new Date(), DateUtil.DATE_FORMAT_NOW);
			Map<String, Object> beans = new HashMap<>();
			beans.put("hasData", 1);
			beans.put("shop", this.shop);
			beans.put("printDate", printDate);
			beans.put("cycleCount", this.cycleCount);
			beans.put("list", this.lstCCMapProduct.getLstObject());

			String folder = ServletActionContext.getServletContext().getRealPath("/") + Configuration.getExcelTemplatePathStock(); //Configuration.getStoreRealPath();
			String templateFileName = folder + ConstantManager.EXPORT_CATEGORY_TEMPLATE;
			templateFileName = templateFileName.replace('/', File.separatorChar);

			String outputName = Configuration.getResourceString(ConstantManager.VI_LANGUAGE, ConstantManager.EXPORT_CATEGORY_STOCK) + this.genExportFileSuffix() + ConstantManager.EXPORT_FILE_EXTENSION;
			String exportFileName = (folder + outputName).replace('/', File.separatorChar);

			inputStream = new BufferedInputStream(new FileInputStream(templateFileName));
			XLSTransformer transformer = new XLSTransformer();
			resultWorkbook = transformer.transformXLS(inputStream, beans);
			os = new BufferedOutputStream(new FileOutputStream(exportFileName));
			resultWorkbook.write(os);
			os.flush();
			this.result.put(Action.ERROR, false);
			String outputPath = ServletActionContext.getServletContext().getContextPath() + Configuration.getExcelTemplatePathStock() + outputName;
			this.result.put(AbstractAction.LIST, outputPath);

		} catch (Exception e) {
			this.result.put(Action.ERROR, true);
			this.result.put("errMsg", Configuration.getResourceString(ConstantManager.VI_LANGUAGE, "system.error"));
			LogUtility.logError(e, e.getMessage(), this.getClass(), null);
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (Exception e) {
					LogUtility.logError(e, e.getMessage(), this.getClass(), null);
				}
			}
			if (os != null) {
				try {
					os.close();
				} catch (Exception e) {
					LogUtility.logError(e, e.getMessage(), this.getClass(), null);
				}
			}
			if (resultWorkbook != null) {
				try {
					resultWorkbook.close();
				} catch (Exception e) {
					LogUtility.logError(e, e.getMessage(), this.getClass(), null);
				}
			}
		}
		return AbstractAction.JSON;
	}

	@Override
	public Staff getStaff() {
		return this.staff;
	}

	@Override
	public void setStaff(Staff staff) {
		this.staff = staff;
	}

	public Shop getShop() {
		return this.shop;
	}

	public void setShop(Shop shop) {
		this.shop = shop;
	}

	public String getStartDate() {
		return this.startDate;
	}

	public void setStartDate(String startDate) {
		this.startDate = startDate;
	}

	public Integer getCycleType() {
		return this.cycleType;
	}

	public void setCycleType(Integer cycleType) {
		this.cycleType = cycleType;
	}

	@Override
	public String getShopCode() {
		return this.shopCode;
	}

	@Override
	public void setShopCode(String shopCode) {
		this.shopCode = shopCode;
	}

	@Override
	public Long getShopId() {
		return this.shopId;
	}

	@Override
	public void setShopId(Long shopId) {
		this.shopId = shopId;
	}

	public String getCycleCountCode() {
		return this.cycleCountCode;
	}

	public void setCycleCountCode(String cycleCountCode) {
		this.cycleCountCode = cycleCountCode;
	}

	public CycleCount getCycleCount() {
		return this.cycleCount;
	}

	public void setCycleCount(CycleCount cycleCount) {
		this.cycleCount = cycleCount;
	}

	public Integer getSortBy() {
		return this.sortBy;
	}

	public void setSortBy(Integer sortBy) {
		this.sortBy = sortBy;
	}

	public Integer getFirstNumber() {
		return this.firstNumber;
	}

	public void setFirstNumber(Integer firstNumber) {
		this.firstNumber = firstNumber;
	}

	public List<ProductInfo> getLstCategoryType() {
		return this.lstCategoryType;
	}

	public void setLstCategoryType(List<ProductInfo> lstCategoryType) {
		this.lstCategoryType = lstCategoryType;
	}

	public Long getCategory() {
		return this.category;
	}

	public void setCategory(Long category) {
		this.category = category;
	}

	public Long getSub_category() {
		return this.sub_category;
	}

	public void setSub_category(Long sub_category) {
		this.sub_category = sub_category;
	}

	public String getProductCode() {
		return this.productCode;
	}

	public void setProductCode(String productCode) {
		this.productCode = productCode;
	}

	public String getProductName() {
		return this.productName;
	}

	public void setProductName(String productName) {
		this.productName = productName;
	}

	public List<ProductInfo> getLstSubCategoryType() {
		return this.lstSubCategoryType;
	}

	public void setLstSubCategoryType(List<ProductInfo> lstSubCategoryType) {
		this.lstSubCategoryType = lstSubCategoryType;
	}

	public void setLstProduct(List<String> lstProduct) {
		this.lstProduct = lstProduct;
	}

	public List<String> getLstProduct() {
		return this.lstProduct;
	}

	public ObjectVO<CycleCountMapProduct> getLstCCMapProduct() {
		return this.lstCCMapProduct;
	}

	public void setLstCCMapProduct(ObjectVO<CycleCountMapProduct> lstCCMapProduct) {
		this.lstCCMapProduct = lstCCMapProduct;
	}

	public void setCycleMapProductId(Long cycleMapProductId) {
		this.cycleMapProductId = cycleMapProductId;
	}

	public Long getCycleMapProductId() {
		return this.cycleMapProductId;
	}

	public String getCycleCountDes() {
		return this.cycleCountDes;
	}

	public Integer getStatus() {
		return this.status;
	}

	public void setStatus(Integer status) {
		this.status = status;
	}

	/**
	 * Sets the cycle count des.
	 *
	 * @param cycleCountDes
	 *            the new cycle count des
	 */
	public void setCycleCountDes(String cycleCountDes) {
		this.cycleCountDes = cycleCountDes;
	}

	public List<Integer> getLstIsDelete() {
		return this.lstIsDelete;
	}

	public void setLstIsDelete(List<Integer> lstIsDelete) {
		this.lstIsDelete = lstIsDelete;
	}

	public List<Integer> getLstStockCardNumber() {
		return this.lstStockCardNumber;
	}

	public void setLstStockCardNumber(List<Integer> lstStockCardNumber) {
		this.lstStockCardNumber = lstStockCardNumber;
	}

	public Integer getLastNumber() {
		return this.lastNumber;
	}

	public void setLastNumber(Integer lastNumber) {
		this.lastNumber = lastNumber;
	}

	public Integer getNoPaging() {
		return this.noPaging;
	}

	public void setNoPaging(Integer noPaging) {
		this.noPaging = noPaging;
	}

	public String getListProductEx() {
		return this.listProductEx;
	}

	public void setListProductEx(String listProductEx) {
		this.listProductEx = listProductEx;
	}

	public Integer getFromAmnt() {
		return this.fromAmnt;
	}

	public void setFromAmnt(Integer fromAmnt) {
		this.fromAmnt = fromAmnt;
	}

	public Integer getToAmnt() {
		return this.toAmnt;
	}

	public void setToAmnt(Integer toAmnt) {
		this.toAmnt = toAmnt;
	}

}
