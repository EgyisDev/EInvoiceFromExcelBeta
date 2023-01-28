using AutoMapper;
using EInvoice.Application.EInvoice.DTOs;

namespace EInvoice.API.Common.Mappers
{
    public class MappingProfile : Profile
    {
        public MappingProfile()
        {
            CreateMap<DocumentRequestDocumentForSerializeDto, DocumentRequestDocumentForSubmitDto>();
        }
    }
}
